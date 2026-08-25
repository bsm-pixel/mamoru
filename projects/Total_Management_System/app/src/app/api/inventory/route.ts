import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/**
 * GET /api/inventory — 재고 현황 집계
 * ?category=가위&search=&low_stock=true
 */
export async function GET(req: NextRequest) {
  const supabase = createServiceClient();
  const sp = req.nextUrl.searchParams;
  const category = sp.get('category') || '';
  const search = sp.get('search') || '';
  const lowStock = sp.get('low_stock') === 'true';
  const lowStockThreshold = parseInt(sp.get('threshold') || '3');

  // 1) 제품 목록 + stock_quantity
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let query = (supabase as any)
    // 112: location_id + 로케이션 코드/라벨 조인 (표시 전용 — 재고 수량 계산에는 일절 관여 안 함)
    .from('products')
    .select('id, name, sku, category, price, price_dealer, price_purchase, price_groups, stock_quantity, raw_stock, is_active, barcode, location_id, warehouse_locations(code, label)')
    .eq('is_active', true)
    .order('name');

  if (category) {
    query = query.eq('category', category);
  }
  if (search) {
    query = query.or(`name.ilike.%${search}%,sku.ilike.%${search}%,barcode.ilike.%${search}%`);
  }

  const { data: products, error: prodErr } = await query;
  if (prodErr) {
    return NextResponse.json({ error: prodErr.message }, { status: 500 });
  }

  // 2) 미입고 수량 (발주완료/선납완료 상태 PO)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data: pendingRaw } = await (supabase as any)
    .from('purchase_order_items')
    .select('product_id, quantity, purchase_orders!inner(status)')
    .in('purchase_orders.status', ['ordered', 'deposit_paid']);

  const pendingMap: Record<string, number> = {};
  if (pendingRaw) {
    for (const row of pendingRaw) {
      if (row.product_id) {
        pendingMap[row.product_id] = (pendingMap[row.product_id] || 0) + row.quantity;
      }
    }
  }

  // 3) 시리얼별 창고 구분 집계 (in_stock만)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data: zoneRaw } = await (supabase as any)
    .from('product_serials')
    .select('product_id, warehouse_zone')
    .eq('status', 'in_stock');

  const zoneMap: Record<string, { raw: number; ready: number; display: number }> = {};
  if (zoneRaw) {
    for (const row of zoneRaw) {
      if (!zoneMap[row.product_id]) {
        zoneMap[row.product_id] = { raw: 0, ready: 0, display: 0 };
      }
      const zone = row.warehouse_zone as 'raw' | 'ready' | 'display';
      if (zone === 'ready') zoneMap[row.product_id].ready++;
      else if (zone === 'display') zoneMap[row.product_id].display++;
      else zoneMap[row.product_id].raw++; // raw 또는 기타
    }
  }

  // 3.5) 반품창고 재고 (status='returned') — 판매가능 현재고와 분리 집계(검수대기·아임웹 미반영)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data: returnedRaw } = await (supabase as any)
    .from('product_serials')
    .select('product_id')
    .eq('status', 'returned');
  const returnMap: Record<string, number> = {};
  if (returnedRaw) {
    for (const row of returnedRaw) returnMap[row.product_id] = (returnMap[row.product_id] || 0) + 1;
  }

  // 비시리얼 반품창고(products.return_stock) — 마이그133 전이면 컬럼 없음 → 에러 무시(빈 맵)
  const returnStockMap: Record<string, number> = {};
  {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data: rs, error: rsErr } = await (supabase as any).from('products').select('id, return_stock');
    if (!rsErr && rs) for (const p of rs) { if (p.return_stock) returnStockMap[p.id] = p.return_stock; }
  }

  // 4) 조합
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const items = (products || []).map((p: any) => ({
    id: p.id,
    name: p.name,
    sku: p.sku,
    category: p.category,
    price: p.price,
    price_dealer: p.price_dealer,
    price_purchase: p.price_purchase,
    barcode: p.barcode,
    stock_quantity: p.stock_quantity || 0,
    raw_stock: p.raw_stock || 0,
    pending_quantity: pendingMap[p.id] || 0,
    zone_raw: p.raw_stock || 0,  // 보관 = raw_stock (비시리얼 수량)
    zone_ready: zoneMap[p.id]?.ready || 0,
    zone_display: zoneMap[p.id]?.display || 0,
    zone_return: (returnMap[p.id] || 0) + (returnStockMap[p.id] || 0),  // 반품창고(시리얼 returned + 비시리얼 return_stock) — 판매가능 현재고 제외
    // 112: 정위치 (표시 전용). 미지정이면 null
    location_id: p.location_id || null,
    location_code: p.warehouse_locations?.code || null,
    location_label: p.warehouse_locations?.label || null,
  }));

  // 저재고 필터
  const filtered = lowStock
    ? items.filter((i: { stock_quantity: number }) => i.stock_quantity <= lowStockThreshold)
    : items;

  // 전체 요약
  const summary = {
    total_products: items.length,
    total_stock: items.reduce((s: number, i: { stock_quantity: number }) => s + i.stock_quantity, 0),
    total_pending: items.reduce((s: number, i: { pending_quantity: number }) => s + i.pending_quantity, 0),
    low_stock_count: items.filter((i: { stock_quantity: number }) => i.stock_quantity <= lowStockThreshold).length,
    total_value: items.reduce((s: number, i: { stock_quantity: number; price_purchase: number }) => s + i.stock_quantity * (i.price_purchase || 0), 0),
  };

  return NextResponse.json({ items: filtered, summary });
}
