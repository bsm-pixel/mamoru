import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

/**
 * POST /api/orders/[id]/exchange — 아임웹 주문 제품 교환 (매출/카드 불변, 상품·재고만 스왑)
 *
 * 원칙:
 *   · 아임웹 주문·결제금액·매출은 절대 건드리지 않음(카드 재결제 방지). 아임웹 재고 push 안 함.
 *   · 반납품 → 반품창고(판매가능 미복원). 시리얼=returned/return zone, 비시리얼=products.return_stock+1.
 *   · 새 제품 = "진짜 −1"(주문 시리얼배정의 raw+1 오프셋 아님 — 원 주문에 없던 제품이라 사전차감 없음).
 *       - serial_ids(모달이 재고/자동생성으로 확보한 in_stock) → in_stock→sold. raw 불변 → stock_quantity −1 자연감소.
 *       - serial 없는 비시리얼 → raw_stock −qty.
 *   · 차액 → cash_transactions '주문교환차액'(현금 등). 주문 status/결제 불변, exchanged_at·memo 기록.
 *
 * body: {
 *   returns:   [{ product_id, product_name?, qty, serial_ids?: string[] }],   // 반납 → 반품창고
 *   new_items: [{ product_id, product_name?, qty, serial_ids?: string[] }],   // 새 제품 → 출고
 *   recovery_method?: string,  // 직접수거|방문수거|택배수거|고객반납
 *   diff_amount?: number,      // + 추가수령 / - 환불 (신제품합 − 주문결제액)
 *   diff_method?: string,      // 현금|카드|이체|없음
 *   memo?: string,
 * }
 */
type ReturnLine = { product_id: string; product_name?: string; qty: number; serial_ids?: string[] };
type NewLine = { product_id: string; product_name?: string; qty: number; serial_ids?: string[] };

export async function POST(request: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id: orderId } = await params;

  const auth = await createServerSupabaseClient();
  const { data: { user } } = await auth.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  let body: {
    returns?: ReturnLine[]; new_items?: NewLine[];
    recovery_method?: string; diff_amount?: number; diff_method?: string; memo?: string;
  };
  try { body = await request.json(); } catch { return NextResponse.json({ error: '잘못된 요청' }, { status: 400 }); }

  const returns = Array.isArray(body.returns) ? body.returns : [];
  const newItems = Array.isArray(body.new_items) ? body.new_items : [];
  const recovery = (body.recovery_method || '직접수거').trim();
  const diffAmount = Number(body.diff_amount || 0);
  const diffMethod = (body.diff_method || '없음').trim();
  if (returns.length === 0 && newItems.length === 0) {
    return NextResponse.json({ error: '반납 또는 새 제품 중 하나는 있어야 합니다' }, { status: 400 });
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db: any = createServiceClient();

  const { data: order, error: oErr } = await db
    .from('orders')
    .select('id, orderer_name, orderer_phone, status, exchange_memo')
    .eq('id', orderId).single();
  if (oErr || !order) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
  if (order.status === 'cancelled') return NextResponse.json({ error: '취소된 주문은 교환할 수 없습니다' }, { status: 400 });

  const now = new Date().toISOString();
  const today = now.slice(0, 10);
  const retLabels: string[] = [];
  const newLabels: string[] = [];

  try {
    // ── 1) 반납품 → 반품창고 (판매가능 미복원) ─────────────────────
    for (const r of returns) {
      if (!r.product_id || !r.qty) continue;
      const serialIds = (r.serial_ids || []).filter(Boolean);

      if (serialIds.length > 0) {
        // 시리얼 반납: 실물에 판매 시리얼 각인 → 반품창고(검수대기·판매불가) 격리
        const { data: srows } = await db.from('product_serials')
          .select('id, warehouse_zone').in('id', serialIds);
        for (const s of (srows || []) as Array<{ id: string; warehouse_zone: string | null }>) {
          await db.from('product_serials').update({
            status: 'returned',
            warehouse_zone: 'return',
            previous_zone: s.warehouse_zone,
            order_id: null, sold_via: null,
          }).eq('id', s.id);
        }
      }
      // 반품창고 카운터 +qty (시리얼/비시리얼 공통 — 물리 실물이 반품창고로 들어옴)
      const { data: p } = await db.from('products').select('return_stock, sku, name').eq('id', r.product_id).single();
      if (p) {
        await db.from('products').update({ return_stock: ((p.return_stock as number) || 0) + r.qty, updated_at: now }).eq('id', r.product_id);
        retLabels.push(`${r.product_name || p.sku || p.name || r.product_id}×${r.qty}`);
      }
      // 판매가능 stock_quantity·raw_stock 는 복원하지 않음 (반품창고行이라 판매불가)
    }

    // ── 2) 새 제품 → 출고 (진짜 −1) ────────────────────────────────
    for (const n of newItems) {
      if (!n.product_id || !n.qty) continue;
      const serialIds = (n.serial_ids || []).filter(Boolean);

      if (serialIds.length > 0) {
        // in_stock → sold (주문 귀속). raw 불변 → stock_quantity 자연 −1. 낙관적 잠금.
        for (const sid of serialIds) {
          const { count } = await db.from('product_serials').update({
            status: 'sold', sold_via: 'online', order_id: orderId,
            offline_sale_id: null, sale_item_id: null, sold_at: now,
            sold_to_name: order.orderer_name || null, sold_to_phone: order.orderer_phone || null,
          }).eq('id', sid).eq('status', 'in_stock')
            .select('id', { count: 'exact', head: true });
          if (count === 0) {
            return NextResponse.json({ error: `시리얼(${sid})이 재고 상태가 아닙니다 — 새로고침 후 다시 시도하세요` }, { status: 409 });
          }
        }
        const { data: p } = await db.from('products').select('sku, name').eq('id', n.product_id).single();
        newLabels.push(`${n.product_name || p?.sku || p?.name || n.product_id}×${serialIds.length}`);
      } else {
        // 비시리얼: 보관(raw)에서 진짜 차감
        const { data: p } = await db.from('products').select('raw_stock, stock_quantity, sku, name').eq('id', n.product_id).single();
        if (!p) return NextResponse.json({ error: `제품(${n.product_id})을 찾을 수 없습니다` }, { status: 404 });
        if ((p.raw_stock || 0) < n.qty) {
          return NextResponse.json({ error: `${p.sku || p.name || '제품'} 보관재고 부족(보관 ${p.raw_stock || 0} < ${n.qty})` }, { status: 400 });
        }
        await db.from('products').update({
          raw_stock: (p.raw_stock || 0) - n.qty,
          stock_quantity: Math.max(0, (p.stock_quantity || 0) - n.qty),
          updated_at: now,
        }).eq('id', n.product_id);
        newLabels.push(`${n.product_name || p.sku || p.name}×${n.qty}`);
      }
    }

    // ── 3) 차액 → cash_transactions ────────────────────────────────
    if (diffAmount !== 0 && diffMethod !== '없음') {
      await db.from('cash_transactions').insert({
        transaction_date: today,
        type: diffAmount > 0 ? 'income' : 'expense',
        category: '주문교환차액',
        amount: Math.abs(diffAmount),
        memo: `주문 ${orderId.slice(0, 8)} 교환차액 ${diffAmount > 0 ? '추가수령' : '환불'}(${diffMethod}): ${order.orderer_name || ''}`.trim(),
        source_type: 'order',
        source_id: orderId,
        created_by: user.id,
      });
    }

    // ── 4) 주문에 교환 이력 기록 (매출/카드/status 불변) ───────────
    const summary = `[교환 ${today}] 반납: ${retLabels.join(', ') || '없음'}→반품창고 / 발송: ${newLabels.join(', ') || '없음'}`
      + ` / 회수: ${recovery}`
      + (diffAmount !== 0 ? ` / 차액 ${diffAmount > 0 ? '+' : '-'}${Math.abs(diffAmount).toLocaleString()} ${diffMethod}` : '')
      + (body.memo ? ` / ${body.memo}` : '');
    const merged = order.exchange_memo ? `${order.exchange_memo}\n${summary}` : summary;
    await db.from('orders').update({ exchanged_at: now, exchange_memo: merged, updated_at: now }).eq('id', orderId);

    return NextResponse.json({ ok: true, summary, returns: retLabels, newItems: newLabels });
  } catch (err) {
    console.error('[orders/exchange]', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
