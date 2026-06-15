/**
 * EVENT 입금확인 → 판매(offline_sales) 자동 전환.
 * 전환 후 발송/배송완료/후기는 기존 판매 인프라가 처리(여기서는 매출·재고만 확정).
 */
import { SLICING_ADDON } from './options';
import type { EventItem } from './types';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function generateSaleNumber(db: any): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `OS-${today}-`;
  const { data } = await db
    .from('offline_sales')
    .select('sale_number')
    .like('sale_number', `${prefix}%`)
    .order('sale_number', { ascending: false })
    .limit(1);
  let seq = 1;
  if (data && data.length > 0) {
    seq = parseInt((data[0].sale_number as string).split('-').pop() || '0', 10) + 1;
  }
  return `${prefix}${String(seq).padStart(3, '0')}`;
}

interface EventRow {
  id: string;
  event_number: string;
  customer_id: string | null;
  customer_name: string;
  customer_phone: string | null;
  items: EventItem[];
  total_amount: number;
}

/** offline_sales + offline_sale_items 생성, 재고 차감, sale_id 반환 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function convertEventToSale(db: any, ev: EventRow): Promise<string> {
  const saleNumber = await generateSaleNumber(db);
  const today = new Date().toISOString().slice(0, 10);

  const { data: sale, error: saleErr } = await db
    .from('offline_sales')
    .insert({
      sale_number: saleNumber,
      customer_id: ev.customer_id,
      customer_name: ev.customer_name,
      customer_phone: ev.customer_phone,
      sale_date: today,
      total_amount: ev.total_amount,
      paid_amount: ev.total_amount,
      payment_method: 'transfer',
      payment_status: 'paid',
      sale_channel: 'offline',
      memo: `EVENT 전환 (${ev.event_number})`,
    })
    .select('id')
    .single();
  if (saleErr || !sale) throw saleErr || new Error('판매 생성 실패');

  // 품목 라인 (제품 sku 조회)
  const productIds = ev.items.map((it) => it.product_id).filter(Boolean) as string[];
  const skuMap: Record<string, string> = {};
  if (productIds.length > 0) {
    const { data: prods } = await db.from('products').select('id, sku').in('id', productIds);
    (prods || []).forEach((p: { id: string; sku: string }) => { skuMap[p.id] = p.sku; });
  }

  const lines = ev.items.map((it) => {
    const qty = Math.max(1, it.qty || 1);
    const unit = it.unit_price || 0;
    return {
      sale_id: sale.id,
      product_id: it.product_id,
      product_name: it.product_name,
      sku: it.product_id ? (skuMap[it.product_id] || null) : null,
      quantity: qty,
      unit_price: unit,
      total_price: unit * qty,
      category: 'EVENT',
    };
  });

  // 슬라이싱 가공 추가비 (별도 라인, 매출/회계 가시성)
  const slicingQty = ev.items.reduce((s, it) => s + (it.slicing ? Math.max(1, it.qty || 1) : 0), 0);
  if (slicingQty > 0) {
    lines.push({
      sale_id: sale.id,
      product_id: null,
      product_name: '슬라이싱 가공',
      sku: null,
      quantity: slicingQty,
      unit_price: SLICING_ADDON,
      total_price: SLICING_ADDON * slicingQty,
      category: 'EVENT',
    });
  }

  const { error: itemsErr } = await db.from('offline_sale_items').insert(lines);
  if (itemsErr) throw itemsErr;

  // 재고 차감 (product_id 있는 라인만, stock_quantity)
  const qtyMap: Record<string, number> = {};
  ev.items.forEach((it) => {
    if (it.product_id) qtyMap[it.product_id] = (qtyMap[it.product_id] || 0) + Math.max(1, it.qty || 1);
  });
  await Promise.all(Object.entries(qtyMap).map(async ([pid, qty]) => {
    const { data: prod } = await db.from('products').select('stock_quantity').eq('id', pid).single();
    if (prod) {
      await db.from('products').update({
        stock_quantity: Math.max(0, (prod.stock_quantity || 0) - qty),
      }).eq('id', pid);
    }
  }));

  return sale.id as string;
}
