/**
 * 복원수리 비용 계산 (GAS L548~L562 로직 이식)
 * 설정값 연동 예정 — 현재는 기본값 사용
 */

/** 수리 단가 */
export const REPAIR_PRICE = {
  mamoru: 10000,   // 마모루 1자루
  other: 20000,    // 타사 1자루
} as const;

/** 배송비 계산 */
export function calcShippingFee(totalQty: number, proceedType: string | null): number {
  if (proceedType === '방문수거') {
    if (totalQty >= 3) return 0;
    if (totalQty === 2) return 3000;
    return 6000; // 1개
  }
  if (proceedType === '직접발송') {
    if (totalQty >= 2) return 0;
    return 3000; // 1개
  }
  return 0; // 매장방문 등
}

/** 서비스 비용 계산 */
export function calcServiceCost(qtyMamoru: number, qtyOther: number): number {
  return qtyMamoru * REPAIR_PRICE.mamoru + qtyOther * REPAIR_PRICE.other;
}

/** 전체 비용 계산 */
export function calcTotalCost(
  qtyMamoru: number,
  qtyOther: number,
  proceedType: string | null
): { serviceCost: number; shippingFee: number; totalAmount: number } {
  const serviceCost = calcServiceCost(qtyMamoru, qtyOther);
  const shippingFee = calcShippingFee(qtyMamoru + qtyOther, proceedType);
  return {
    serviceCost,
    shippingFee,
    totalAmount: serviceCost + shippingFee,
  };
}
