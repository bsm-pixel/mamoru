import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

/**
 * POST /api/orders/[id]/serials — 아임웹 주문에 시리얼 배정/해제 (무결성 핵심)
 *
 * body: { serialIds: string[] }  = 이 주문에 배정할 "재고(in_stock) 시리얼 id" 최종 집합
 *   현재 배정과 비교해 추가분(assign)·제거분(unassign)을 산출해 반영한다.
 *
 * ⚠️ 재고 불변식: stock_quantity = raw_stock + (in_stock 시리얼 수)
 *   아임웹 주문은 sync 시 이미 stock_quantity·raw_stock 둘 다 −qty 차감된 상태(loose 가정).
 *   시리얼 배정은 "loose 차감분"을 "시리얼 소진"으로 사후 치환하는 것이므로:
 *     · assign  → 시리얼 in_stock→sold,  raw_stock +1   (stock_quantity 불변)
 *     · unassign→ 시리얼 sold→in_stock,  raw_stock −1   (stock_quantity 불변)
 *   → 좌변 불변, 우변 상쇄. stock_quantity 재차감·아임웹 재고 재push 절대 금지(이중차감).
 */
export async function POST(request: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  const { id: orderId } = await params;

  // 인증
  const auth = await createServerSupabaseClient();
  const { data: { user } } = await auth.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  let body: { serialIds?: unknown; manualByProduct?: unknown };
  try { body = await request.json(); } catch { body = {}; }
  const target = Array.isArray(body.serialIds) ? body.serialIds.filter((x): x is string => typeof x === 'string') : [];
  // 수동/자동생성 시리얼 번호 (product_id → 번호[])
  const manualByProduct: Record<string, string[]> = {};
  if (body.manualByProduct && typeof body.manualByProduct === 'object') {
    for (const [pid, arr] of Object.entries(body.manualByProduct as Record<string, unknown>)) {
      if (Array.isArray(arr)) {
        const nums = arr.filter((x): x is string => typeof x === 'string' && x.trim() !== '').map((x) => x.trim());
        if (nums.length) manualByProduct[pid] = nums;
      }
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db: any = createServiceClient();

  // 주문 + 품목(수량 상한 검증용) + 현재 배정 시리얼 로드
  const [orderRes, itemsRes, currentRes] = await Promise.all([
    db.from('orders').select('id, orderer_name, orderer_phone').eq('id', orderId).single(),
    db.from('order_items').select('imweb_product_no, product_id, quantity').eq('order_id', orderId),
    db.from('product_serials').select('id, product_id, warehouse_zone, previous_zone').eq('order_id', orderId).eq('status', 'sold'),
  ]);
  if (orderRes.error || !orderRes.data) return NextResponse.json({ error: '주문을 찾을 수 없습니다' }, { status: 404 });
  const order = orderRes.data as { id: string; orderer_name: string | null; orderer_phone: string | null };

  const current = (currentRes.data || []) as Array<{ id: string; product_id: string | null; warehouse_zone: string | null; previous_zone: string | null }>;
  const currentIds = new Set(current.map((s) => s.id));
  const targetSet = new Set(target);

  const toAddIds = target.filter((sid) => !currentIds.has(sid));
  const toRemove = current.filter((s) => !targetSet.has(s.id));

  // 추가 대상 시리얼 로드(in_stock 확인 + product_id 파악)
  let toAdd: Array<{ id: string; product_id: string | null; warehouse_zone: string | null }> = [];
  if (toAddIds.length > 0) {
    const { data: addRows } = await db
      .from('product_serials')
      .select('id, product_id, warehouse_zone')
      .in('id', toAddIds)
      .eq('status', 'in_stock');
    toAdd = (addRows || []) as typeof toAdd;
    if (toAdd.length !== toAddIds.length) {
      return NextResponse.json(
        { error: `배정 대상 ${toAddIds.length}개 중 ${toAdd.length}개만 재고 상태 (이미 판매/배정된 시리얼 확인)` },
        { status: 409 }
      );
    }
  }

  // 수량 상한 검증: 제품별 (현재 배정 − 제거 + 추가) ≤ 주문 수량
  // 구(舊) 품목은 product_id 가 비어있을 수 있어 imweb_product_no 로 보완 해석.
  const orderItems = (itemsRes.data || []) as Array<{ imweb_product_no: string | null; product_id: string | null; quantity: number }>;
  const missingNos = [...new Set(orderItems.filter((i) => !i.product_id && i.imweb_product_no).map((i) => i.imweb_product_no as string))];
  const noToPid: Record<string, string> = {};
  if (missingNos.length > 0) {
    const { data: prods } = await db.from('products').select('id, imweb_product_no').in('imweb_product_no', missingNos);
    (prods || []).forEach((p: { id: string; imweb_product_no: string | null }) => { if (p.imweb_product_no) noToPid[String(p.imweb_product_no)] = p.id; });
  }
  const orderQtyByProduct: Record<string, number> = {};
  for (const it of orderItems) {
    const pid = it.product_id || (it.imweb_product_no ? noToPid[String(it.imweb_product_no)] : null);
    if (pid) orderQtyByProduct[pid] = (orderQtyByProduct[pid] || 0) + (it.quantity || 0);
  }
  const finalCountByProduct: Record<string, number> = {};
  const countProduct = (pid: string | null, d: number) => { if (pid) finalCountByProduct[pid] = (finalCountByProduct[pid] || 0) + d; };
  current.forEach((s) => countProduct(s.product_id, 1));
  toRemove.forEach((s) => countProduct(s.product_id, -1));
  toAdd.forEach((s) => countProduct(s.product_id, 1));
  Object.entries(manualByProduct).forEach(([pid, nums]) => countProduct(pid, nums.length));
  for (const [pid, cnt] of Object.entries(finalCountByProduct)) {
    const cap = orderQtyByProduct[pid] ?? 0;
    if (cnt > cap) {
      return NextResponse.json({ error: `제품 시리얼 배정 수(${cnt})가 주문 수량(${cap})을 초과합니다` }, { status: 400 });
    }
  }

  const now = new Date().toISOString();
  const rawDeltaByProduct: Record<string, number> = {};
  const addRaw = (pid: string | null, d: number) => { if (pid) rawDeltaByProduct[pid] = (rawDeltaByProduct[pid] || 0) + d; };

  // 1) 추가: in_stock → sold (낙관적 잠금) + 주문 귀속. raw_stock +1
  for (const s of toAdd) {
    const { count } = await db
      .from('product_serials')
      .update({
        previous_zone: s.warehouse_zone,
        status: 'sold',
        sold_via: 'online',
        order_id: orderId,
        offline_sale_id: null,
        sale_item_id: null,
        sold_at: now,
        sold_to_name: order.orderer_name || null,
        sold_to_phone: order.orderer_phone || null,
      })
      .eq('id', s.id)
      .eq('status', 'in_stock')
      .select('id', { count: 'exact', head: true });
    if (count === 0) {
      return NextResponse.json({ error: `시리얼 ${s.id}이(가) 방금 다른 곳에 배정되었습니다. 새로고침 후 다시 시도하세요` }, { status: 409 });
    }
    addRaw(s.product_id, +1);
  }

  // 2) 제거: sold(이 주문) → in_stock + 귀속 해제. raw_stock −1
  for (const s of toRemove) {
    await db
      .from('product_serials')
      .update({
        status: 'in_stock',
        warehouse_zone: s.previous_zone || s.warehouse_zone || 'ready',
        sold_via: null,
        order_id: null,
        sold_at: null,
        sold_to_name: null,
        sold_to_phone: null,
      })
      .eq('id', s.id)
      .eq('order_id', orderId);
    addRaw(s.product_id, -1);
  }

  // 2.5) 수동/자동생성 시리얼: 재고에 없어 새 번호를 만드는 경우.
  //   · 신규 생성 = sold 라벨만. raw 변화 없음 (주문이 이미 loose로 raw −1 했고, 이 라벨이 그 loose 단위를 가리킴).
  //   · 기존 in_stock 번호를 입력한 것이면 = 재고 배정과 동일 (raw +1).
  //   · 타 판매/주문/계약 소속이면 차단(이전 동의 미지원 MVP).
  let manualAdded = 0;
  for (const [pid, numbers] of Object.entries(manualByProduct)) {
    for (const num of numbers) {
      const { data: existRows } = await db
        .from('product_serials')
        .select('id, status, order_id, warehouse_zone')
        .eq('serial_number', num)
        .limit(1);
      const ex = existRows && existRows[0];
      if (ex) {
        if (ex.status === 'sold' && ex.order_id === orderId) continue; // 이미 이 주문에 배정됨
        if (ex.status === 'in_stock') {
          const { count } = await db
            .from('product_serials')
            .update({
              previous_zone: ex.warehouse_zone, status: 'sold', sold_via: 'online', order_id: orderId,
              offline_sale_id: null, sale_item_id: null, sold_at: now,
              sold_to_name: order.orderer_name || null, sold_to_phone: order.orderer_phone || null,
            })
            .eq('id', ex.id).eq('status', 'in_stock')
            .select('id', { count: 'exact', head: true });
          if (count === 0) return NextResponse.json({ error: `시리얼 "${num}" 배정 충돌 — 새로고침 후 재시도` }, { status: 409 });
          addRaw(pid, +1);
          manualAdded++;
        } else {
          return NextResponse.json({ error: `시리얼 "${num}" 은(는) 이미 다른 판매/주문에 등록되어 있습니다` }, { status: 409 });
        }
      } else {
        await db.from('product_serials').insert({
          serial_number: num, product_id: pid, status: 'sold', warehouse_zone: 'ready', previous_zone: 'ready',
          sold_via: 'online', order_id: orderId, sold_at: now,
          sold_to_name: order.orderer_name || null, sold_to_phone: order.orderer_phone || null,
        });
        manualAdded++;
      }
    }
  }

  // 3) raw_stock 치환 반영 (stock_quantity·아임웹 재고는 손대지 않음)
  await Promise.all(Object.entries(rawDeltaByProduct).map(async ([pid, delta]) => {
    if (delta === 0) return;
    const { data: prod } = await db.from('products').select('raw_stock').eq('id', pid).single();
    if (!prod) return;
    const next = Math.max(0, (prod.raw_stock || 0) + delta);
    await db.from('products').update({ raw_stock: next, updated_at: now }).eq('id', pid);
  }));

  // 최신 배정 목록 반환
  const { data: after } = await db
    .from('product_serials')
    .select('id, product_id, serial_number, status, sold_at')
    .eq('order_id', orderId);

  return NextResponse.json({ ok: true, serials: after || [], added: toAdd.length + manualAdded, removed: toRemove.length });
}
