/**
 * 납품 합계 계산 — 복원수리(category='RS') 항목은 VAT 대상에서 제외.
 *
 * - 제품 항목에만 부가세 적용, RS 금액은 그대로 총액에 더함(공급가로 분류).
 * - 할인은 **제품 → 복원수리 순서로 소진**한다(2026-09-15 fix).
 *   제품 금액을 먼저 깎고, 남은 할인은 RS 금액에서 뺀다. VAT 는 제품분에만 적용되므로 RS 할인은 세액에 영향 없음.
 *   (기존엔 제품에만 적용 → 복원수리 전용 납품에 할인을 넣으면 할인이 통째로 사라져 total_amount 가 할인 전 금액으로 저장됐다:
 *    DL-20260914-001 복원수리 9 + 배송비(RS) − 할인 2만 = 7.3만이어야 하는데 9.3만 저장 · 미수금도 9.3만)
 *
 * 하위호환(기존 동작과 동일):
 *  - 제품 전용(RS 0): 기존 (itemTotal - discount) 기준과 수치 동일
 *  - RS 전용(제품 0, vat='none'): 총액 = RS합 − 할인, vat = 0 (할인 0이면 기존과 동일)
 *  - 혼합(제품 + RS): 제품만 VAT, RS는 무세 가산 (신규, 올바른 처리)
 */
export interface DeliveryTotalItem {
  category?: string | null;
  quantity: number;
  unit_price: number;
}

export interface DeliveryTotals {
  supplyAmount: number;
  vatAmount: number;
  totalAmount: number;
}

export function computeDeliveryTotals(
  items: DeliveryTotalItem[],
  vatType: 'included' | 'separate' | 'none',
  discount: number,
): DeliveryTotals {
  const rsTotal = items
    .filter((i) => i.category === 'RS')
    .reduce((s, i) => s + i.quantity * i.unit_price, 0);
  const productTotal = items
    .filter((i) => i.category !== 'RS')
    .reduce((s, i) => s + i.quantity * i.unit_price, 0);
  const discountVal = discount || 0;
  const productBase = Math.max(0, productTotal - discountVal);
  // 제품에서 다 못 깎은 할인은 복원수리 금액에서 소진 (복원수리 전용 납품 할인 유실 방지)
  const rsBase = Math.max(0, rsTotal - Math.max(0, discountVal - productTotal));

  let pSupply = 0, pVat = 0, pTotal = 0;
  if (vatType === 'separate') {
    pSupply = productBase;
    pVat = Math.round(productBase * 0.1);
    pTotal = productBase + pVat;
  } else if (vatType === 'none') {
    pSupply = productBase;
    pVat = 0;
    pTotal = productBase;
  } else { // included
    pSupply = Math.round(productBase / 1.1);
    pVat = productBase - pSupply;
    pTotal = productBase;
  }

  return {
    supplyAmount: pSupply + rsBase,
    vatAmount: pVat,
    totalAmount: pTotal + rsBase,
  };
}
