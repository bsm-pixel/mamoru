/**
 * EVENT 가격 계산 (서버 권위). 고객 폼(page_form.html JS)에도 동일 로직이 복제돼 있으니 함께 수정할 것.
 * 규칙: 같은 단가끼리 묶음 → min_qty 도달 시 묶음 반복 + 나머지 정가. 슬라이싱 가공은 별도 가산.
 */
import { SLICING_ADDON } from './options';
import type { EventItem, DiscountRule } from './types';

export interface EventPricing {
  base: number;       // 묶음 할인 적용 후 품목 합계
  slicing: number;    // 슬라이싱 가공 합계
  total: number;      // base + slicing
  listTotal: number;  // 할인 전 정가 합계(품목) — 표시용
  discount: number;   // listTotal - base
}

export function computeEventPricing(items: EventItem[], rules: DiscountRule[] = []): EventPricing {
  const byPrice: Record<number, number> = {};
  let slicing = 0;
  let listTotal = 0;
  for (const it of items) {
    const qty = Math.max(1, Number(it.qty) || 1);
    const price = Math.max(0, Number(it.unit_price) || 0);
    byPrice[price] = (byPrice[price] || 0) + qty;
    listTotal += price * qty;
    if (it.slicing) slicing += SLICING_ADDON * qty;
  }
  let base = 0;
  for (const [priceStr, qty] of Object.entries(byPrice)) {
    const price = Number(priceStr);
    const rule = (rules || []).find(
      (r) => Number(r.unit_price) === price && r.min_qty > 0 && r.bundle_price > 0 && qty >= r.min_qty,
    );
    if (rule) {
      const bundles = Math.floor(qty / rule.min_qty);
      const rem = qty % rule.min_qty;
      base += bundles * rule.bundle_price + rem * price;
    } else {
      base += qty * price;
    }
  }
  return { base, slicing, total: base + slicing, listTotal, discount: listTotal - base };
}
