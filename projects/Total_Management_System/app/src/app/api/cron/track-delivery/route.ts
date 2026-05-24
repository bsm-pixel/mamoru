import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { queryStatus } from '@/lib/lotte/client';
import { queryTrackingStatus } from '@/lib/lotte/alps-client';

/**
 * GET /api/cron/track-delivery
 *
 * 1) 아임웹 orders: shipping → delivered (queryStatus, 기존)
 * 2) 복원수리 repairs: shipped → delivered (queryTrackingStatus = ALPS '91' 인수자등록, 2026-05-24 추가)
 *
 * → 사장님이 ALPS 직접 확인 없이 TMS 카드 status 로 배송완료 즉시 확인 가능
 */
export async function GET(request: NextRequest) {
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const url = new URL(request.url);
  const debug = url.searchParams.get('debug') === '1';
  const debugResults: Array<{ as_id: string; invoice: string; state: string; detail?: string }> = [];

  try {
    // 🚨 cron 은 user 인증 없으므로 service role 클라이언트 필수 (RLS 우회)
    //    2026-05-24: createServerSupabaseClient (cookie 기반) 쓰던 버그 발견 → RLS 막혀 0건 처리됨
    const supabase = createServiceClient();

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // [1] 아임웹 orders 추적 (shipping → delivered)
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    const { data: orders, error: ordersErr } = await (supabase as any)
      .from('orders')
      .select('id, invoice_number, imweb_order_no, orderer_name, review_requested_at')
      .eq('status', 'shipping')
      .not('invoice_number', 'is', null)
      .limit(50);

    if (ordersErr) throw ordersErr;

    let ordersDelivered = 0;
    for (const order of orders || []) {
      try {
        const result = await queryStatus(order.invoice_number);
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

    return NextResponse.json({
      orders: { checked: orders?.length || 0, delivered: ordersDelivered },
      repairs: { checked: repairs?.length || 0, delivered: repairsDelivered },
      ...(debug && {
        debug: {
          env: {
            LOTTE_TRACK_API_URL: process.env.LOTTE_TRACK_API_URL ? `set:${process.env.LOTTE_TRACK_API_URL.length}chars` : 'MISSING',
            LOTTE_CLIENT_KEY:    process.env.LOTTE_CLIENT_KEY    ? `set:${process.env.LOTTE_CLIENT_KEY.length}chars`    : 'MISSING',
            LOTTE_JOBCUSTCD:     (process.env.LOTTE_JOB_CUST_CD || process.env.LOTTE_JOBCUSTCD) ? `set:${(process.env.LOTTE_JOB_CUST_CD || process.env.LOTTE_JOBCUSTCD || '').length}chars` : 'MISSING',
          },
          // 첫 3건만 상세 노출 (보안 + payload 크기)
          firstResults: debugResults.slice(0, 3),
        },
      }),
    });
  } catch (err) {
    console.error('[cron/track-delivery] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
