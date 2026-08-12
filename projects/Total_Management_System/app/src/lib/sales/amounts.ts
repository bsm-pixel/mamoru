/**
 * 금액 공식 SSOT — 판매/납품의 net·미수금·완납실수납 단일 진실점 (2026-08-12)
 *
 * ⚠️ 두 도메인은 total_amount 컨벤션이 반대다. 반드시 도메인별 헬퍼를 쓸 것.
 *   - offline_sales : total_amount = GROSS(정가합)   → net = total - discount
 *   - deliveries    : total_amount = NET(할인 이미 반영, computeDeliveryTotals) → net = total (할인 재차감 금지)
 *
 * 배경: 같은 공식이 30여 곳에 인라인 복붙되며 discount 누락/이중차감 버그가 반복됨.
 *   (판매 미수금 할인 누락 56631b22, 납품 할인 이중차감 잠재). 이후 모든 금액 계산은 이 헬퍼 경유.
 */

type SaleLike = { total_amount?: number | null; discount_amount?: number | null; paid_amount?: number | null };
type DeliveryLike = { total_amount?: number | null; paid_amount?: number | null };

/** offline_sales 순매출 = 소계(gross) - 할인 */
export function saleNet(r: SaleLike): number {
  return (r.total_amount || 0) - (r.discount_amount || 0);
}

/** offline_sales 미수금 = max(0, net - 실수납) */
export function saleOutstanding(r: SaleLike): number {
  return Math.max(0, saleNet(r) - (r.paid_amount || 0));
}

/** offline_sales 완납 시 실수납액 = net (할인 반영) */
export function saleFullPaid(r: SaleLike): number {
  return Math.max(0, saleNet(r));
}

/** deliveries 순매출 = total (할인은 total_amount에 이미 반영됨 — 재차감 금지) */
export function deliveryNet(r: DeliveryLike): number {
  return r.total_amount || 0;
}

/** deliveries 미수금 = max(0, total - 실수납) */
export function deliveryOutstanding(r: DeliveryLike): number {
  return Math.max(0, (r.total_amount || 0) - (r.paid_amount || 0));
}

/** deliveries 완납 시 실수납액 = total (이미 net) */
export function deliveryFullPaid(r: DeliveryLike): number {
  return Math.max(0, r.total_amount || 0);
}
