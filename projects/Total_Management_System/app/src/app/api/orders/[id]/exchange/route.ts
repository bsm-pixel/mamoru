import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { insertReturn } from '@/lib/returns/insert-return';

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
    recovery_method?: string; ship_method?: string; diff_amount?: number; diff_method?: string; memo?: string;
  };
  try { body = await request.json(); } catch { return NextResponse.json({ error: '잘못된 요청' }, { status: 400 }); }

  const returns = Array.isArray(body.returns) ? body.returns : [];
  const newItems = Array.isArray(body.new_items) ? body.new_items : [];
  const recovery = (body.recovery_method || '직접수거').trim();
  const shipMethod = (body.ship_method || '배송').trim();  // 배송|직접전달 (새 제품 발송 방식)
  const diffAmount = Number(body.diff_amount || 0);
  const diffMethod = (body.diff_method || '없음').trim();
  if (returns.length === 0 && newItems.length === 0) {
    return NextResponse.json({ error: '반납 또는 새 제품 중 하나는 있어야 합니다' }, { status: 400 });
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db: any = createServiceClient();

  const { data: order, error: oErr } = await db
    .from('orders')
    .select('id, orderer_name, orderer_phone, customer_id, recipient_postcode, recipient_address, recipient_address_detail, status, exchange_memo')
    .eq('id', orderId).single();
  if (oErr || !order) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
  if (order.status === 'cancelled') return NextResponse.json({ error: '취소된 주문은 교환할 수 없습니다' }, { status: 400 });

  const now = new Date().toISOString();
  const today = now.slice(0, 10);
  const retLabels: string[] = [];
  const newLabels: string[] = [];

  try {
    // ── 1) 반납품 → 반품창고 (판매가능 미복원, 멱등) ─────────────────
    //   🚨 서버에서 '이 주문·이 제품의 판매됨 시리얼'을 직접 조회해 반품 처리.
    //   교환을 두 번 돌려도(재실행) 이미 반품된 건 0건 조회 → return_stock 이중집계 없음.
    for (const r of returns) {
      if (!r.product_id || !r.qty) continue;

      const { data: soldRows } = await db.from('product_serials')
        .select('id, warehouse_zone')
        .eq('order_id', orderId).eq('product_id', r.product_id).eq('status', 'sold');
      const sold = (soldRows || []) as Array<{ id: string; warehouse_zone: string | null }>;

      // 이 제품이 이 주문에서 시리얼 추적되는지(과거 반품분 포함) — 비시리얼 판별
      const { count: anySerial } = await db.from('product_serials')
        .select('id', { count: 'exact', head: true })
        .eq('order_id', orderId).eq('product_id', r.product_id);
      const isSerialTracked = (anySerial || 0) > 0;

      let moved = 0;
      for (const s of sold) {
        const { count } = await db.from('product_serials').update({
          status: 'returned', warehouse_zone: 'return', previous_zone: s.warehouse_zone,
          // order_id 유지(감사·멱등 근거). sold_via 유지.
        }).eq('id', s.id).eq('status', 'sold').select('id', { count: 'exact', head: true });
        if ((count || 0) > 0) moved++;
      }

      // 반품창고 증가: 시리얼 제품이면 '이번에 실제 이동한 수'만(멱등), 비시리얼이면 qty
      const addReturn = isSerialTracked ? moved : r.qty;
      const { data: p } = await db.from('products').select('return_stock, sku, name').eq('id', r.product_id).single();
      if (p) {
        if (addReturn > 0) {
          await db.from('products').update({ return_stock: ((p.return_stock as number) || 0) + addReturn, updated_at: now }).eq('id', r.product_id);
        }
        retLabels.push(`${r.product_name || p.sku || p.name || r.product_id}×${r.qty}`);
      }
      // 판매가능(stock_quantity·raw_stock)은 복원 안 함 — 반품창고行(검수/복원 대상)
    }

    // ── 2) 새 제품 → 출고 (진짜 −1) ────────────────────────────────
    for (const n of newItems) {
      if (!n.product_id || !n.qty) continue;
      const serialIds = (n.serial_ids || []).filter(Boolean);

      if (serialIds.length > 0) {
        // in_stock → sold (주문 귀속). 낙관적 잠금.
        let soldCount = 0;
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
          soldCount++;
        }
        // 🚨 시리얼 in_stock→sold = 판매가능 재고 소진 → stock_quantity −soldCount (raw 불변).
        //   stock_quantity는 저장 컬럼이라 명시 차감 필요(빠뜨리면 현재고가 +1씩 부풀어 드리프트).
        const { data: p } = await db.from('products').select('stock_quantity, sku, name').eq('id', n.product_id).single();
        if (p) {
          await db.from('products').update({
            stock_quantity: Math.max(0, (p.stock_quantity || 0) - soldCount), updated_at: now,
          }).eq('id', n.product_id);
        }
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
    const exchangeGoods = newLabels.map((l) => l.replace(/×\d+$/, '')).join(', ');  // 송장 품목명(×N 제거)
    const summary = `[교환 ${today}] 반납: ${retLabels.join(', ') || '없음'}→반품창고 / 발송: ${newLabels.join(', ') || '없음'}(${shipMethod})`
      + ` / 회수: ${recovery}`
      + (diffAmount !== 0 ? ` / 차액 ${diffAmount > 0 ? '+' : '-'}${Math.abs(diffAmount).toLocaleString()} ${diffMethod}` : '')
      + (body.memo ? ` / ${body.memo}` : '');
    const merged = order.exchange_memo ? `${order.exchange_memo}\n${summary}` : summary;
    await db.from('orders').update({
      exchanged_at: now, exchange_memo: merged,
      exchange_ship_method: shipMethod, exchange_goods: exchangeGoods || null,
      updated_at: now,
    }).eq('id', orderId);

    // ── 5) returns SSOT 레코드 생성 (교환·반품 화면에서 관리) — 주문당 1건, 재실행 시 중복 방지 ──
    try {
      const { data: existing } = await db.from('returns').select('id').eq('order_id', orderId).limit(1);
      if (!existing || existing.length === 0) {
        await insertReturn(db, today, {
          return_type: 'exchange',
          source: 'order',
          order_id: orderId,
          sale_id: null,
          product_id: returns[0]?.product_id || null,
          product_name: retLabels.join(', ') || null,          // 반납품 요약
          new_product_id: newItems[0]?.product_id || null,
          new_product_name: newLabels.map((l) => l.replace(/×\d+$/, '')).join(', ') || null,  // 발송품 요약
          customer_id: order.customer_id || null,
          name: order.orderer_name || null,
          phone: order.orderer_phone || null,
          postcode: order.recipient_postcode || null,
          address: order.recipient_address || null,
          address_detail: order.recipient_address_detail || null,
          pickup_method: recovery,
          status: 'inbound',        // 구제품 반품창고 확정(수거완료). 남은 건 발송/마감.
          inbound_at: now,
          reason: '주문 교환',
          memo: summary,
          created_by: user.id,
        });
      }
    } catch (e) { console.error('[orders/exchange] returns 레코드 생성 실패(교환은 완료):', e); }

    return NextResponse.json({ ok: true, summary, returns: retLabels, newItems: newLabels, shipMethod });
  } catch (err) {
    console.error('[orders/exchange]', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
