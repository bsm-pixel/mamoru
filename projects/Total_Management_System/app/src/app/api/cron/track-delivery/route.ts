import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { queryStatus } from '@/lib/lotte/client';
import { queryTrackingStatus } from '@/lib/lotte/alps-client';
import { sendReviewRequestNotification } from '@/lib/notification/review-request';
import { getServerSetting } from '@/hooks/use-settings';

/**
 * GET /api/cron/track-delivery
 *
 * 1) 아임웹 orders: shipping → delivered (queryStatus)
 * 2) 복원수리 repairs: shipped → delivered (queryTrackingStatus = ALPS 41/45 인수자등록)
 * 3) [예정] offline_sales: shipped_at→delivered_at (Phase 2)
 *
 * 2026-05-25 추가:
 *   - orders 자동 후기요청 발송 (settings 토글 ON + 가드 통과 시)
 *   - debug=1 시 ordersFirstResults 노출
 *
 * → 사장님이 ALPS 직접 확인 없이 TMS 카드 status 로 배송완료 즉시 확인 가능
 */
export async function GET(request: NextRequest) {
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const url = new URL(request.url);
  // ?debug=1 : 운영 진단 도구 (CRON_SECRET 인증 필요)
  //   응답에 LOTTE 환경변수 길이/cus path + 첫 3건 ALPS 결과 노출
  //   사장님 운영 중 자동 추적 의문 발생 시 즉시 진단 가능
  //   (2026-05-25 client.ts jobCustCd fallback 버그 발견에 활용 — 백성민 케이스)
  const debug = url.searchParams.get('debug') === '1';
  const debugResults: Array<{ as_id: string; invoice: string; state: string; detail?: string }> = [];
  const ordersDebugResults: Array<{ order_no: string; invoice: string; state: string; raw_code?: string }> = [];

  try {
    // 🚨 cron 은 user 인증 없으므로 service role 클라이언트 필수 (RLS 우회)
    //    2026-05-24: createServerSupabaseClient (cookie 기반) 쓰던 버그 발견 → RLS 막혀 0건 처리됨
    const supabase = createServiceClient();

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [1] 아임웹 orders 추적 (shipping → delivered) + 자동 후기요청 (2026-05-25 확장)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const { data: orders, error: ordersErr } = await (supabase as any)
      .from('orders')
      .select('id, invoice_number, imweb_order_no, orderer_name, orderer_phone, review_requested_at')
      .eq('status', 'shipping')
      .not('invoice_number', 'is', null)
      .limit(50);

    if (ordersErr) throw ordersErr;

    let ordersDelivered = 0;
    for (const order of orders || []) {
      try {
        const result = await queryStatus(order.invoice_number);
        if (debug) {
          ordersDebugResults.push({
            order_no: order.imweb_order_no,
            invoice: order.invoice_number,
            state: result.state,
            raw_code: String((result.raw as Record<string, unknown> | undefined)?.code || ''),
          });
        }
        if (result.ok && result.state === 'DELIVERED') {
          await (supabase as any)
            .from('orders')
            .update({
              status: 'delivered',
              delivered_at: new Date().toISOString(),
            })
            .eq('id', order.id);
          ordersDelivered++;
          console.log(`[track-delivery/orders] ${order.imweb_order_no} → 배송완료`);

          // 후기요청 자동 발송 (settings 토글 ON + 중복 방지)
          // 2026-05-26 메모: 사장님 정책 "약속 받은 고객만 자동 발송" 적용 불가 — orders 에 review_promised_at 컬럼 미존재
          //   아임웹 무인 주문 흐름이라 사장님이 약속 받을 일 거의 없음 → 현재 정책 유지 (토글 ON 시 모든 배송완료 자동 발송)
          //   향후 사장님 정책 통일 원하면 orders.review_promised_at 컬럼 마이그레이션 필요
          after(async () => {
            try {
              const autoEnabled = await getServerSetting<boolean>(supabase as any, 'review.auto_request_on_completion', false);
              if (!autoEnabled) {
                console.log(`[track-delivery/orders auto-review] ${order.imweb_order_no} skip — 토글 OFF`);
                return;
              }
              if (order.review_requested_at) {
                console.log(`[track-delivery/orders auto-review] ${order.imweb_order_no} skip — 이미 발송됨`);
                return;
              }
              if (!order.orderer_phone) return;

              const r = await sendReviewRequestNotification({
                source: 'sale',
                sourceId: order.imweb_order_no,
                customerName: order.orderer_name || '고객',
                customerPhone: order.orderer_phone,
                reviewType: 'purchase',
              });
              if (r.success) {
                await (supabase as any).from('orders')
                  .update({ review_requested_at: new Date().toISOString() })
                  .eq('id', order.id);
                console.log(`[track-delivery/orders auto-review] ${order.imweb_order_no} 발송 성공`);
              } else {
                console.error(`[track-delivery/orders auto-review] 실패:`, r.error);
              }
            } catch (e) {
              console.error(`[track-delivery/orders auto-review] 예외:`, e);
            }
          });
        }
      } catch (e) {
        console.error(`[track-delivery/orders] ${order.imweb_order_no} 추적 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [2] 복원수리 repairs 추적 (shipped → delivered) — 2026-05-24 추가
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   - status='shipped' + invoice_number 있는 건 50건 폴링
    //   - ALPS queryTrackingStatus '91' (인수자등록) 감지 시 자동 delivered 전환
    //   - 합포장 출고 케이스(다른 주문 송장 복사된 건)도 동일 흐름 (invoice_number 채워져 있음)
    const { data: repairs, error: repairsErr } = await (supabase as any)
      .from('repairs')
      .select('id, as_id, invoice_number, name')
      .eq('status', 'shipped')
      .not('invoice_number', 'is', null)
      .limit(50);

    if (repairsErr) throw repairsErr;

    let repairsDelivered = 0;
    for (const repair of repairs || []) {
      try {
        const result = await queryTrackingStatus(repair.invoice_number);
        if (debug) {
          debugResults.push({
            as_id: repair.as_id,
            invoice: repair.invoice_number,
            state: result.state,
            detail: result.detail,
          });
        }
        if (result.state === 'DELIVERED') {
          await (supabase as any)
            .from('repairs')
            .update({
              status: 'delivered',
              delivered_at: new Date().toISOString(),
            })
            .eq('id', repair.id);

          // 이력 기록 (자동 전환임을 명시)
          await (supabase as any).from('repair_history').insert({
            repair_id: repair.id,
            from_status: 'shipped',
            to_status: 'delivered',
            changed_by: null, // 자동 cron
            note: 'ALPS 추적 자동 감지 (인수자등록)',
          });

          repairsDelivered++;
          console.log(`[track-delivery/repairs] ${repair.as_id} (${repair.name}) → 배송완료`);
        }
      } catch (e) {
        console.error(`[track-delivery/repairs] ${repair.as_id} 추적 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [3] offline_sales 추적 (shipped → delivered) — 2026-05-25 추가 (Phase 2)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   진입 조건: 송장 발급됨(invoice_number) + 출고완료(shipped_at) + 배송 미완료(delivered_at NULL) + 미취소
    //   매장 직접 수령은 invoice_number NULL → 자연 제외
    const { data: sales, error: salesErr } = await (supabase as any)
      .from('offline_sales')
      .select('id, sale_number, invoice_number, customer_name, customer_phone, review_requested_at, review_promised_at, review_promised_type, review_promised_subtype')
      .not('shipped_at', 'is', null)
      .is('delivered_at', null)
      .not('invoice_number', 'is', null)
      .is('cancelled_at', null)
      .limit(50);

    if (salesErr) throw salesErr;

    let salesDelivered = 0;
    for (const sale of sales || []) {
      try {
        const result = await queryTrackingStatus(sale.invoice_number);
        if (result.state === 'DELIVERED') {
          await (supabase as any)
            .from('offline_sales')
            .update({ delivered_at: new Date().toISOString() })
            .eq('id', sale.id);
          salesDelivered++;
          console.log(`[track-delivery/offline_sales] ${sale.sale_number} (${sale.customer_name}) → 배송완료`);

          // 자동 후기요청 (2026-05-26 정책 정정: 약속 ✓ 고객만 자동 발송 — 보수적)
          //   사장님 의도: "약속 받은 고객 + 배송완료(인수자등록) 자동 감지 → 자동 알림톡"
          //   약속 X 고객은 사장님 수동 발송만 (compact UI 의 후기 요청 버튼 사용)
          after(async () => {
            try {
              const autoEnabled = await getServerSetting<boolean>(supabase as any, 'review.auto_request_on_completion', false);
              if (!autoEnabled) {
                console.log(`[track-delivery/offline_sales auto-review] ${sale.sale_number} skip — 토글 OFF`);
                return;
              }
              if (!sale.review_promised_at) {
                console.log(`[track-delivery/offline_sales auto-review] ${sale.sale_number} skip — 약속 X (사장님 수동만)`);
                return;
              }
              if (sale.review_requested_at) {
                console.log(`[track-delivery/offline_sales auto-review] ${sale.sale_number} skip — 이미 발송됨`);
                return;
              }
              if (!sale.customer_phone) return;

              // 094: 약속 시 선택한 유형으로 발송 (NULL이면 'purchase' 디폴트 — 백필 데이터 호환)
              // 095: subtype 도 함께 전달 (purchase 면 무시)
              const reviewType = (sale.review_promised_type as 'purchase' | 'repair' | 'consult' | null) || 'purchase';
              const subtype = reviewType === 'purchase' ? undefined : (sale.review_promised_subtype as string | null) || undefined;
              const r = await sendReviewRequestNotification({
                source: 'sale',
                sourceId: sale.sale_number,
                customerName: sale.customer_name || '고객',
                customerPhone: sale.customer_phone,
                reviewType,
                subtype,
              });
              if (r.success) {
                await (supabase as any).from('offline_sales')
                  .update({ review_requested_at: new Date().toISOString() })
                  .eq('id', sale.id);
                console.log(`[track-delivery/offline_sales auto-review] ${sale.sale_number} 발송 성공`);
              } else {
                console.error(`[track-delivery/offline_sales auto-review] 실패:`, r.error);
              }
            } catch (e) {
              console.error(`[track-delivery/offline_sales auto-review] 예외:`, e);
            }
          });
        }
      } catch (e) {
        console.error(`[track-delivery/offline_sales] ${sale.sale_number} 추적 실패:`, e);
      }
    }

    return NextResponse.json({
      orders: { checked: orders?.length || 0, delivered: ordersDelivered },
      repairs: { checked: repairs?.length || 0, delivered: repairsDelivered },
      sales: { checked: sales?.length || 0, delivered: salesDelivered },
      ...(debug && {
        debug: {
          env: (() => {
            const trackUrl = process.env.LOTTE_TRACK_API_URL || '';
            // URL 의 cus/XXX 부분만 추출해서 노출 (전체 URL 노출 아님, endpoint 식별용)
            const cusMatch = trackUrl.match(/cus\/([^/]+)\//);
            return {
              LOTTE_TRACK_API_URL_length: trackUrl.length,
              LOTTE_TRACK_API_URL_cus: cusMatch ? cusMatch[1] : 'NOT_MATCHED',
              LOTTE_CLIENT_KEY_length: (process.env.LOTTE_CLIENT_KEY || '').length,
              LOTTE_JOBCUSTCD_length: (process.env.LOTTE_JOB_CUST_CD || process.env.LOTTE_JOBCUSTCD || '').length,
            };
          })(),
          // 첫 3건만 상세 노출 (보안 + payload 크기)
          firstResults: debugResults.slice(0, 3),
          ordersFirstResults: ordersDebugResults.slice(0, 3),  // 2026-05-25 추가
        },
      }),
    });
  } catch (err) {
    console.error('[cron/track-delivery] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
