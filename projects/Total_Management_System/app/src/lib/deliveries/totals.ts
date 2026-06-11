/**
 * 납품 합계 계산 — 복원수리(category='RS') 항목은 VAT 대상에서 제외.
 *
 * - 제품 항목에만 부가세 적용, RS 금액은 그대로 총액에 더함(공급가로 분류).
 * - 할인은 제품 기준에 적용(복원수리는 할인 미적용 정책 — repair UI에 할인 입력 없음).
 *
 * 하위호환(기존 동작과 동일):
 *  - 제품 전용(RS 0): 기존 (itemTotal - discount) 기준과 수치 동일
 *  - RS 전용(제품 0, vat='none'): 총액 = RS합, vat = 0 — 기존과 동일
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
  const productBase = Math.max(0, productTotal - (discount || 0));

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
    supplyAmount: pSupply + rsTotal,
    vatAmount: pVat,
    totalAmount: pTotal + rsTotal,
  };
}
