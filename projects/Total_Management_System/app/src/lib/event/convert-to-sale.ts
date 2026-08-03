/**
 * EVENT 입금확인 → 판매(offline_sales) 자동 전환.
 * 전환 후 발송/배송완료/후기는 기존 판매 인프라가 처리(여기서는 매출·재고만 확정).
 */
import { SLICING_ADDON } from './options';
import { updateImwebStock } from '@/lib/imweb/client';
import type { EventItem } from './types';
import { insertOfflineSale } from '@/lib/sales/insert-offline-sale';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';

interface EventRow {
  id: string;
  event_number: string;
  customer_id: string | null;
  customer_name: string;
  customer_phone: string | null;
  receive_method?: string | null;   // visit | delivery — 주소 세팅 판단용
  postcode?: string | null;
  address1?: string | null;
  address2?: string | null;
  items: EventItem[];
  total_amount: number;
}

/** 접수 → 판매 전환 옵션 (재고판매 등 EVENT 외 종류가 재사용) */
export interface ConvertOptions {
  /** offline_sale_items.category 값 (기본 'EVENT') */
  category?: string;
  /** memo 접두 라벨 (기본 'EVENT 전환') */
  memoLabel?: string;
}

/** offline_sales + offline_sale_items 생성, 재고 차감, sale_id 반환 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function convertEventToSale(db: any, ev: EventRow, opts?: ConvertOptions): Promise<string> {
  const category = opts?.category ?? 'EVENT';
  const memoLabel = opts?.memoLabel ?? 'EVENT 전환';
  const today = new Date().toISOString().slice(0, 10);

  // 정가 합계(품목 라인 합 = 단가×수량 + 슬라이싱) → 묶음 할인 = 정가합 − 최종금액
  const slicingQtyAll = ev.items.reduce((s, it) => s + (it.slicing ? Math.max(1, it.qty || 1) : 0), 0);
  const lineSum = ev.items.reduce((s, it) => s + (it.unit_price || 0) * Math.max(1, it.qty || 1), 0)
    + SLICING_ADDON * slicingQtyAll;
  const discountAmount = Math.max(0, lineSum - ev.total_amount);

  // 🔗 고객 연결 보장 — 접수 시 customer_id 가 비어있으면(전화 없이 접수됐거나 매칭 실패) 여기서 전화/이름으로 매칭·생성.
  //    안 하면 판매가 '게스트'로 저장돼 송장 생성(customer_id 필수)·미수금 집계 등이 막힌다. (2026-08-02)
  let customerId = ev.customer_id;
  if (!customerId && ev.customer_phone) {
    const isVisit = ev.receive_method === 'visit';
    const { customerId: matched } = await matchOrCreateCustomer(db, {
      phone: ev.customer_phone,
      name: ev.customer_name,
      source: category === 'LS' ? 'stock_sale' : 'event',
      extra: {
        addressRoad: isVisit ? null : (ev.address1 || null),
        addressDetail: isVisit ? null : (ev.address2 || null),
        postcode: isVisit ? null : (ev.postcode || null),
      },
    });
    customerId = matched;
  }

  // 판매번호 채번+중복재시도는 공용 insertOfflineSale (SSOT)
  const sale = await insertOfflineSale(db, today, {
    customer_id: customerId,
    customer_name: ev.customer_name,
    customer_phone: ev.customer_phone,
    sale_date: today,
    total_amount: ev.total_amount,
    discount_amount: discountAmount,
    paid_amount: ev.total_amount,
    payment_method: 'transfer',
    payment_status: 'paid',
    sale_channel: 'offline',
    memo: `${memoLabel} (${ev.event_number})${discountAmount > 0 ? ` · 묶음할인 -${discountAmount.toLocaleString()}` : ''}`,
  });

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
      category,
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

  // 상품 외 가산액 (재고판매 배송비 등) — total 이 품목합보다 크면 그 차액을 별도 라인으로.
  // 이렇게 해야 판매 합계 = offline_sale_items 합계 가 맞다. (EVENT 는 total<=품목합 이라 미발생)
  const extra = ev.total_amount - lineSum;
  if (extra > 0) {
    lines.push({
      sale_id: sale.id,
      product_id: null,
      product_name: category === 'LS' ? '배송비' : '추가금',
      sku: null,
      quantity: 1,
      unit_price: extra,
      total_price: extra,
      category,
    });
  }

  const { error: itemsErr } = await db.from('offline_sale_items').insert(lines);
  if (itemsErr) throw itemsErr;

  // 재고 차감 — EVENT 품목은 시리얼 없는 판매이므로 stock_quantity + raw_stock(보관) 둘 다 차감
  // (판매 route의 시리얼 없는 판매와 동일 규칙. 보관 미차감 시 현재고<보관 불일치 발생 — 2026-06-16 버그 수정)
  const qtyMap: Record<string, number> = {};
  ev.items.forEach((it) => {
    if (it.product_id) qtyMap[it.product_id] = (qtyMap[it.product_id] || 0) + Math.max(1, it.qty || 1);
  });
  await Promise.all(Object.entries(qtyMap).map(async ([pid, qty]) => {
    const { data: prod } = await db.from('products').select('stock_quantity, raw_stock, imweb_product_no').eq('id', pid).single();
    if (prod) {
      await db.from('products').update({
        stock_quantity: Math.max(0, (prod.stock_quantity || 0) - qty),
        raw_stock: Math.max(0, (prod.raw_stock || 0) - qty),
      }).eq('id', pid);
      if (prod.imweb_product_no) {
        try { await updateImwebStock(Number(prod.imweb_product_no), -qty); }
        catch (e) { console.error('[event/convert] 아임웹 재고 동기화 실패:', prod.imweb_product_no, e); }
      }
    }
  }));

  return sale.id as string;
}
