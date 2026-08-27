import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { queryStatus } from '@/lib/lotte/client';
import { queryTrackingStatus } from '@/lib/lotte/alps-client';
import { sendReviewRequestNotification } from '@/lib/notification/review-request';
import { sendSalesShippedNotification, sendExchangeShippedNotification } from '@/lib/notification/sales-shipped';
import { sendNotification } from '@/lib/notification/make-webhook';
import { isB2BCustomerType } from '@/lib/sales/customer-type';
import { getServerSetting } from '@/hooks/use-settings';
import { shipImwebOrder } from '@/lib/imweb/client';

/**
 * GET /api/cron/track-delivery
 *
 * [1] 아임웹 orders   : shipping → delivered (queryStatus)
 * [2] 복원수리 repairs : shipped → delivered (ALPS 41/45 인수자등록) + 후기요청 (109 추가)
 * [2-A] 복원수리 집하 : ready_to_ship → shipped (ALPS 10 집하) + as_shipped 알림톡  ← 109 신규
 * [3-A] 판매 집하     : 송장O·미출고 → shipped_at 자동 기록 + B2C 출고 알림톡      ← 109 신규
 * [3] 판매 offline_sales: shipped_at → delivered_at + 후기요청
 * [4] B2B 납품 deliveries: → delivered_at
 *
 * 109 (2026-07-12) — 집하 자동 감지:
 *   롯데 기사님이 방문 수거하며 스캔하면 ALPS 가 '집하'(godsStatCd 10)로 바뀐다.
 *   그 코드는 원래부터 응답에 왔지만 09/41/45 만 보느라 버려지고 있었다.
 *   → 사장님의 수동 클릭 2회([출고완료] + 알림톡 체크)를 0회로.
 *   ⚠️ 알림톡은 B2C 만. B2B(딜러·아카데미) 는 발송하지 않는다.
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
  // 109: 집하 감지 진단 — ALPS tracking 레코드의 실제 필드명(trackingKeys) 확인용
  const salesPickupDebug: Array<{
    sale_number: string; invoice: string; state: string;
    pickedUp: boolean; pickedUpAt?: string; trackingKeys?: string[];
  }> = [];

  try {
    // 🚨 cron 은 user 인증 없으므로 service role 클라이언트 필수 (RLS 우회)
    //    2026-05-24: createServerSupabaseClient (cookie 기반) 쓰던 버그 발견 → RLS 막혀 0건 처리됨
    const supabase = createServiceClient();
    // 109: 신규 블록용 별칭 — 기존 코드의 `db` 반복 대신 한 번만 캐스팅 (lint 에러 증가 0)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [1-A] 아임웹 orders 집하 감지 (ready_to_ship → shipping) + 아임웹 배송중(send) back-sync — 128 신규
    //   송장생성 시 status='ready_to_ship'(배송대기). 기사 집하가 감지되면 배송중으로 올리고
    //   아임웹에도 send 로 배송중을 반영한다. (복원수리 [2-A]·납품 [4-A] 와 동일 패턴)
    //   ⚠️ [1] 앞에 둔다: 여기서 shipping 으로 올리면 [1]이 같은 사이클에 배달완료까지 처리 가능.
    const { data: orderPickups, error: orderPickupErr } = await db
      .from('orders')
      .select('id, invoice_number, imweb_order_no')
      .eq('status', 'ready_to_ship')
      .not('invoice_number', 'is', null)
      .limit(50);

    if (orderPickupErr) throw orderPickupErr;

    let ordersShipped = 0;
    for (const order of orderPickups || []) {
      try {
        const r = await queryTrackingStatus(order.invoice_number);
        if (r.state === 'CANCELLED' || r.state === 'NOT_FOUND') continue;
        if (!r.pickedUp) continue;   // 아직 기사님이 안 가져감

        // 조건부 CAS — status 가 아직 ready_to_ship 일 때만 전이 (중복 방지)
        const { data: claimed } = await db
          .from('orders')
          .update({
            status: 'shipping',
            shipped_at: r.pickedUpAt || new Date().toISOString(),
          })
          .eq('id', order.id)
          .eq('status', 'ready_to_ship')
          .select('id');

        if (!claimed || claimed.length === 0) continue;
        ordersShipped++;
        console.log(`[track-delivery/orders pickup] ${order.imweb_order_no} → 배송중(집하) 자동 처리`);

        // 아임웹 배송중 back-sync (send) — 전 품목
        if (order.imweb_order_no) {
          after(async () => {
            try {
              await shipImwebOrder(order.imweb_order_no);
            } catch (e) {
              console.error(`[track-delivery/orders pickup] ${order.imweb_order_no} 아임웹 send 실패:`, e);
            }
          });
        }
      } catch (e) {
        console.error(`[track-delivery/orders pickup] ${order.imweb_order_no} 집하 확인 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [1] 아임웹 orders 추적 (shipping → delivered) + 자동 후기요청 (2026-05-25 확장)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const { data: orders, error: ordersErr } = await db
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
          await db
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
              const autoEnabled = await getServerSetting<boolean>(db, 'review.auto_request_on_completion', false);
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
              if (r.success && !r.skipped) {
                await db.from('orders')
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
    // ── [2-A] 복원수리 집하 감지 (ready_to_ship → shipped) — 109 신규 ──
    //   송장 발급 시 status='ready_to_ship' (api/repair/[id]/ship). 기사님이 수거하면 자동 출고 처리.
    const { data: repairPickups, error: repairPickupErr } = await db
      .from('repairs')
      .select('id, as_id, invoice_number, name, phone, courier_name')
      .eq('status', 'ready_to_ship')
      .not('invoice_number', 'is', null)
      .limit(50);

    if (repairPickupErr) throw repairPickupErr;

    let repairsPicked = 0;
    for (const repair of repairPickups || []) {
      try {
        const r = await queryTrackingStatus(repair.invoice_number);
        if (r.state === 'CANCELLED' || r.state === 'NOT_FOUND') continue;
        if (!r.pickedUp) continue;

        // 조건부 CAS — status 가 아직 ready_to_ship 일 때만 전이 (중복 알림톡 차단)
        const { data: claimed } = await db
          .from('repairs')
          .update({
            status: 'shipped',
            shipped_at: r.pickedUpAt || new Date().toISOString(),
            shipped_source: 'alps_pickup',
          })
          .eq('id', repair.id)
          .eq('status', 'ready_to_ship')
          .select('id');

        if (!claimed || claimed.length === 0) continue;
        repairsPicked++;

        await db.from('repair_history').insert({
          repair_id: repair.id,
          from_status: 'ready_to_ship',
          to_status: 'shipped',
          changed_by: null,                       // 자동 cron
          note: '롯데 집하 자동 감지 (기사님 수거 스캔)',
        });
        console.log(`[track-delivery/repairs pickup] ${repair.as_id} (${repair.name}) → 출고완료(수거) 자동 처리`);

        // 출고 알림톡 (as_shipped) — 이미 배달완료된 건은 skip (이미 받은 고객에게 "출고했습니다" 금지)
        if (r.state === 'DELIVERED') {
          console.log(`[track-delivery/repairs pickup] ${repair.as_id} 알림톡 skip — 이미 배달완료`);
          continue;
        }
        if (!repair.phone) continue;
        after(async () => {
          try {
            const sent = await sendNotification({
              template: 'as_shipped',
              phone: repair.phone,
              name: repair.name,
              data: {
                // 🔴 id + as_uid 둘 다 필수 — Make 매핑이 어느 키를 읽든 #{as_uid}(수리내역 조회 버튼)가 채워지도록
                id: repair.as_id,
                as_uid: repair.as_id,
                courier: repair.courier_name || '롯데택배',
                tracking: repair.invoice_number || '',
              },
            });
            if (!sent.success) console.error(`[track-delivery/repairs pickup] ${repair.as_id} 알림톡 실패:`, sent.error);
            else console.log(`[track-delivery/repairs pickup] ${repair.as_id} 출고 알림톡 발송 성공`);
          } catch (e) {
            console.error(`[track-delivery/repairs pickup] ${repair.as_id} 알림톡 예외:`, e);
          }
        });
      } catch (e) {
        console.error(`[track-delivery/repairs pickup] ${repair.as_id} 집하 확인 실패:`, e);
      }
    }

    const { data: repairs, error: repairsErr } = await db
      .from('repairs')
      .select('id, as_id, invoice_number, name, phone, review_promised_at, review_promised_type, review_promised_subtype, review_request_sent_at')
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
          await db
            .from('repairs')
            .update({
              status: 'delivered',
              delivered_at: new Date().toISOString(),
            })
            .eq('id', repair.id);

          // 이력 기록 (자동 전환임을 명시)
          await db.from('repair_history').insert({
            repair_id: repair.id,
            from_status: 'shipped',
            to_status: 'delivered',
            changed_by: null, // 자동 cron
            note: 'ALPS 추적 자동 감지 (인수자등록)',
          });

          repairsDelivered++;
          console.log(`[track-delivery/repairs] ${repair.as_id} (${repair.name}) → 배송완료`);

          // 🔴 109 버그 수정: 자동 후기요청 (판매와 동일 정책 — '약속한 고객만')
          //    기존엔 이 크론이 DB 를 직접 update 해서 PATCH /api/repair/[id] 의 발송 코드를 우회했다.
          //    → 사장님이 수동으로 상태를 바꿀 때만 나가고, ALPS 자동 배송완료 건은 발송 자체가 없었음.
          after(async () => {
            try {
              const autoEnabled = await getServerSetting<boolean>(db, 'review.auto_request_on_completion', false);
              if (!autoEnabled) {
                console.log(`[track-delivery/repairs auto-review] ${repair.as_id} skip — 토글 OFF`);
                return;
              }
              if (!repair.review_promised_at) {
                console.log(`[track-delivery/repairs auto-review] ${repair.as_id} skip — 약속 X (사장님 수동만)`);
                return;
              }
              if (repair.review_request_sent_at) {
                console.log(`[track-delivery/repairs auto-review] ${repair.as_id} skip — 이미 발송됨`);
                return;
              }
              if (!repair.phone) return;

              const reviewType = (repair.review_promised_type as 'purchase' | 'repair' | 'consult' | null) || 'repair';
              const subtype = reviewType === 'purchase' ? undefined : (repair.review_promised_subtype as string | null) || undefined;
              const r = await sendReviewRequestNotification({
                source: 'repair',
                sourceId: repair.as_id,
                customerName: repair.name || '고객',
                customerPhone: repair.phone,
                reviewType,
                subtype,
              });
              if (r.success && !r.skipped) {
                await db.from('repairs')
                  .update({ review_request_sent_at: new Date().toISOString() })
                  .eq('id', repair.id);
                console.log(`[track-delivery/repairs auto-review] ${repair.as_id} 발송 성공`);
              } else {
                console.error(`[track-delivery/repairs auto-review] 실패:`, r.error);
              }
            } catch (e) {
              console.error(`[track-delivery/repairs auto-review] 예외:`, e);
            }
          });
        }
      } catch (e) {
        console.error(`[track-delivery/repairs] ${repair.as_id} 추적 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [3-A] offline_sales 집하 감지 (송장O·미출고 → 자동 출고완료 + B2C 출고 알림톡) — 109 신규
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   진입 조건: 송장 있음 + 아직 출고 전 + 배송 미완료 + 미취소 + 최근 30일
    //   ⚠️ 30일 필터: "송장만 뽑고 영영 안 나가는 건"(취소 누락 등)이 limit 50 을 영구 점유해
    //      신건이 밀리는 적체를 막는다. 30일 넘은 미출고 건은 데이터 오류 → 사장님 수동 정리.
    //   ⚠️ 반드시 [3] 앞에 둔다: 여기서 shipped_at 을 채우면 [3] 의 SELECT 가 그 건을 바로 주워
    //      같은 사이클에 delivered_at 까지 처리한다(집하+배달완료가 함께 잡힌 건).
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
    const { data: pickupTargets, error: pickupErr } = await db
      .from('offline_sales')
      .select('id, sale_number, invoice_number, customer_name, customer_phone, customer_type, courier_name')
      .not('invoice_number', 'is', null)
      .is('shipped_at', null)
      .is('delivered_at', null)
      .is('cancelled_at', null)
      .gte('sale_date', thirtyDaysAgo)
      .order('sale_date', { ascending: false })
      .limit(50);

    if (pickupErr) throw pickupErr;

    let salesPicked = 0;
    let salesNotified = 0;
    let salesPickupSkippedB2B = 0;
    for (const sale of pickupTargets || []) {
      try {
        const r = await queryTrackingStatus(sale.invoice_number);
        if (debug && salesPickupDebug.length < 3) {
          salesPickupDebug.push({
            sale_number: sale.sale_number, invoice: sale.invoice_number,
            state: r.state, pickedUp: !!r.pickedUp, pickedUpAt: r.pickedUpAt,
            trackingKeys: r.trackingKeys,   // ALPS 실제 필드명 확인용 (확정 후 제거 가능)
          });
        }
        // 송장 취소/미접수 건은 건드리지 않는다
        if (r.state === 'CANCELLED' || r.state === 'NOT_FOUND') continue;
        if (!r.pickedUp) continue;   // 아직 기사님이 안 가져감

        // 🔒 조건부 CAS — shipped_at 이 아직 NULL 일 때만 채운다.
        //    크론 중복 실행 / 사장님 수동 버튼 동시 클릭에도 알림톡이 정확히 1회만 나가게 하는 핵심.
        const { data: claimed } = await db
          .from('offline_sales')
          .update({
            shipped_at: r.pickedUpAt || new Date().toISOString(),
            shipped_source: 'alps_pickup',
          })
          .eq('id', sale.id)
          .is('shipped_at', null)
          .select('id');

        if (!claimed || claimed.length === 0) continue;   // 다른 경로가 선점 → 알림톡 skip
        salesPicked++;
        console.log(`[track-delivery/offline_sales pickup] ${sale.sale_number} (${sale.customer_name}) → 출고완료(수거) 자동 처리`);

        // 🔴 이미 배달완료된 건은 출고 알림톡을 보내지 않는다.
        //    집하 시점을 통째로 놓친 케이스(장애·배포 공백·과거 미출고 건).
        //    이미 물건을 받은 고객에게 "출고했습니다"가 가면 혼란만 준다.
        if (r.state === 'DELIVERED') {
          console.log(`[track-delivery/offline_sales pickup] ${sale.sale_number} 알림톡 skip — 이미 배달완료`);
          continue;
        }
        // B2B(딜러·아카데미)는 출고 알림톡 대상 아님
        if (isB2BCustomerType(sale.customer_type)) {
          salesPickupSkippedB2B++;
          console.log(`[track-delivery/offline_sales pickup] ${sale.sale_number} 알림톡 skip — B2B(${sale.customer_type})`);
          continue;
        }

        after(async () => {
          try {
            const sent = await sendSalesShippedNotification(db, {
              id: sale.id,
              saleNumber: sale.sale_number,
              invoiceNumber: sale.invoice_number,
              customerName: sale.customer_name || '고객',
              customerPhone: sale.customer_phone,
              customerType: sale.customer_type,
              courierName: sale.courier_name,
            });
            if (sent.sent) {
              await db.from('offline_sales')
                .update({ shipped_notified_at: new Date().toISOString() })
                .eq('id', sale.id);
              console.log(`[track-delivery/offline_sales pickup] ${sale.sale_number} 출고 알림톡 발송 성공`);
            } else {
              console.log(`[track-delivery/offline_sales pickup] ${sale.sale_number} 알림톡 미발송 — ${sent.reason} ${sent.error || ''}`);
            }
          } catch (e) {
            console.error(`[track-delivery/offline_sales pickup] ${sale.sale_number} 알림톡 예외:`, e);
          }
        });
        salesNotified++;
      } catch (e) {
        console.error(`[track-delivery/offline_sales pickup] ${sale.sale_number} 집하 확인 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [3-C] returns 교환 출고 집하 감지 → 교환 출고 알림톡 (136, 2026-08-27)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   배송 교환건의 '교환 출고 송장'이 집하되면 판매 출고와 동일하게 sales_shipped 로 발송.
    //   매장 교환(송장 없음)은 exchange_out_invoice_number NULL → 자연 제외. B2B 는 헬퍼가 걸러 냄.
    //   dedup: exchange_out_notified_at 을 CAS 로 선점(판매 shipped_at 선점과 동일 철학 — 집하 시 1회).
    let exchangePicked = 0;
    let exchangeNotified = 0;
    const { data: exTargets, error: exErr } = await db
      .from('returns')
      .select('id, return_number, sale_id, new_product_name, exchange_out_invoice_number, exchange_out_courier_name, name, phone')
      .eq('return_type', 'exchange')
      .not('exchange_out_invoice_number', 'is', null)
      .is('exchange_out_notified_at', null)
      .not('status', 'in', '(cancelled)')
      .gte('exchange_shipped_at', thirtyDaysAgo)
      .limit(50);
    // 🛡️ 비치명적: 마이그136(exchange_out_notified_at) 미실행 등으로 실패해도 나머지 배송추적은 계속
    if (exErr) console.error('[track-delivery/returns pickup] 조회 실패(마이그136 미실행?):', exErr.message);

    for (const ret of (exErr ? [] : exTargets) || []) {
      try {
        const r = await queryTrackingStatus(ret.exchange_out_invoice_number);
        // 송장 취소/미접수는 건드리지 않음
        if (r.state === 'CANCELLED' || r.state === 'NOT_FOUND') continue;
        if (!r.pickedUp) continue;   // 아직 기사님이 안 가져감

        // 🔒 조건부 CAS — exchange_out_notified_at 이 NULL 일 때만 선점(중복 발송 방지, 집하 시 1회)
        const { data: claimed } = await db
          .from('returns')
          .update({ exchange_out_notified_at: r.pickedUpAt || new Date().toISOString() })
          .eq('id', ret.id)
          .is('exchange_out_notified_at', null)
          .select('id');
        if (!claimed || claimed.length === 0) continue;   // 다른 사이클이 선점
        exchangePicked++;

        // 이미 배달완료된 건은 "출고했습니다" 금지(집하 시점을 통째로 놓친 케이스)
        if (r.state === 'DELIVERED') {
          console.log(`[track-delivery/returns pickup] ${ret.return_number} 알림톡 skip — 이미 배달완료`);
          continue;
        }

        // 원 판매번호·고객유형(B2B 가드) 조회 — 없으면 반품번호로 표시
        let refId: string | null = ret.return_number;
        let customerType: string | null = null;
        if (ret.sale_id) {
          const { data: s } = await db.from('offline_sales').select('sale_number, customer_type').eq('id', ret.sale_id).single();
          if (s) { refId = s.sale_number || refId; customerType = s.customer_type; }
        }

        after(async () => {
          try {
            const sent = await sendExchangeShippedNotification({
              refId,
              invoiceNumber: ret.exchange_out_invoice_number,
              customerName: ret.name || '고객',
              customerPhone: ret.phone,
              customerType,
              newProductName: ret.new_product_name,
              courierName: ret.exchange_out_courier_name,
            });
            if (sent.sent) console.log(`[track-delivery/returns pickup] ${ret.return_number} 교환 출고 알림톡 발송 성공`);
            else console.log(`[track-delivery/returns pickup] ${ret.return_number} 알림톡 미발송 — ${sent.reason} ${sent.error || ''}`);
          } catch (e) {
            console.error(`[track-delivery/returns pickup] ${ret.return_number} 알림톡 예외:`, e);
          }
        });
        exchangeNotified++;
      } catch (e) {
        console.error(`[track-delivery/returns pickup] ${ret.return_number} 집하 확인 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [3] offline_sales 추적 (shipped → delivered) — 2026-05-25 추가 (Phase 2)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   진입 조건: 송장 발급됨(invoice_number) + 출고완료(shipped_at) + 배송 미완료(delivered_at NULL) + 미취소
    //   매장 직접 수령은 invoice_number NULL → 자연 제외
    const { data: sales, error: salesErr } = await db
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
          await db
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
              const autoEnabled = await getServerSetting<boolean>(db, 'review.auto_request_on_completion', false);
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
              if (r.success && !r.skipped) {
                await db.from('offline_sales')
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

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [4-A] deliveries(B2B 납품) 집하 감지 (출고대기 → 출고완료) — 110 신규 (2026-07-12)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   전엔 api/lotte/book 이 송장 발급과 동시에 status='shipped' 를 강제했다.
    //   → 기사님이 오지도 않았는데 화면에 "출고완료"가 떴다(사장님 지적). 그 강제를 제거하고 여기서 올린다.
    //   B2B 는 출고 알림톡을 보내지 않는다 (사장님 결정) — 상태만 정확히.
    const { data: dlPickups, error: dlPickupErr } = await db
      .from('deliveries')
      .select('id, dl_number, customer_name, tracking_number')
      .eq('status', 'confirmed')
      .not('tracking_number', 'is', null)
      .is('cancelled_at', null)
      .limit(50);

    if (dlPickupErr) throw dlPickupErr;

    let deliveriesPicked = 0;
    for (const dl of dlPickups || []) {
      try {
        const r = await queryTrackingStatus(dl.tracking_number);
        if (r.state === 'CANCELLED' || r.state === 'NOT_FOUND') continue;
        if (!r.pickedUp) continue;   // 아직 기사님이 안 가져감

        // 조건부 CAS — 수동 [출고 완료] 버튼과 동시 실행돼도 한 번만
        const { data: claimed } = await db
          .from('deliveries')
          .update({
            status: 'shipped',
            // shipped_date 는 date 형(YYYY-MM-DD). 집하 시각을 알면 그 날짜, 모르면 오늘.
            shipped_date: (r.pickedUpAt || new Date().toISOString()).slice(0, 10),
            shipped_source: 'alps_pickup',
            updated_at: new Date().toISOString(),
          })
          .eq('id', dl.id)
          .eq('status', 'confirmed')
          .select('id');

        if (!claimed || claimed.length === 0) continue;
        deliveriesPicked++;
        console.log(`[track-delivery/deliveries pickup] ${dl.dl_number} (${dl.customer_name}) → 출고완료(수거) 자동 처리`);
      } catch (e) {
        console.error(`[track-delivery/deliveries pickup] ${dl.dl_number} 집하 확인 실패:`, e);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [4] deliveries(B2B 납품) 추적 → delivered_at — 2026-06-01 추가
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    //   진입 조건: 송장(tracking_number) 있음 + 배송 미완료(delivered_at NULL) + 미취소
    //   ALPS 인수자등록(코드 45)/배달완료(41) 감지 시 delivered_at 세팅 (status 는 정산용이라 미변경)
    const { data: deliveries, error: dlErr } = await db
      .from('deliveries')
      .select('id, dl_number, customer_name, tracking_number')
      .not('tracking_number', 'is', null)
      .is('delivered_at', null)
      .is('cancelled_at', null)
      .limit(50);

    if (dlErr) throw dlErr;

    let deliveriesDelivered = 0;
    for (const dl of deliveries || []) {
      try {
        const result = await queryTrackingStatus(dl.tracking_number);
        if (result.state === 'DELIVERED') {
          await db
            .from('deliveries')
            .update({ delivered_at: new Date().toISOString() })
            .eq('id', dl.id);
          deliveriesDelivered++;
          console.log(`[track-delivery/deliveries] ${dl.dl_number} (${dl.customer_name}) → 배송완료`);
        }
      } catch (e) {
        console.error(`[track-delivery/deliveries] ${dl.dl_number} 추적 실패:`, e);
      }
    }

    return NextResponse.json({
      ordersPickup: { checked: orderPickups?.length || 0, shipped: ordersShipped },   // 128: 주문 집하 감지
      orders: { checked: orders?.length || 0, delivered: ordersDelivered },
      // 109: 집하 자동 감지 (기사님 수거 스캔 → 출고완료)
      repairsPickup: { checked: repairPickups?.length || 0, picked: repairsPicked },
      salesPickup: {
        checked: pickupTargets?.length || 0,
        picked: salesPicked,
        notified: salesNotified,
        skippedB2B: salesPickupSkippedB2B,
      },
      exchangePickup: { checked: exTargets?.length || 0, picked: exchangePicked, notified: exchangeNotified },   // 136: 교환 출고 집하 감지
      deliveriesPickup: { checked: dlPickups?.length || 0, picked: deliveriesPicked },   // 110
      repairs: { checked: repairs?.length || 0, delivered: repairsDelivered },
      sales: { checked: sales?.length || 0, delivered: salesDelivered },
      deliveries: { checked: deliveries?.length || 0, delivered: deliveriesDelivered },
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
          // 109: 집하 감지 진단 — trackingKeys 로 ALPS 실제 날짜 필드명 확인 (확정 후 제거 가능)
          salesPickupFirstResults: salesPickupDebug.slice(0, 3),
        },
      }),
    });
  } catch (err) {
    console.error('[cron/track-delivery] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
