/**
 * 고객 유형 판정 (B2C / B2B) — 단일 출처
 *
 * 마이그 078 정의: B2B = customer_type IN ('dealer','academy')
 *                 B2C = customer_type IS NULL OR NOT IN ('dealer','academy')
 *
 * ⚠️ 같은 판정식이 use-sales.ts / use-dashboard-stats.ts / api/reports/summary 에 흩어져 있다.
 *    새 사본을 만들지 말고 이 함수를 쓸 것 (기존 3곳 통합은 별도 사이클).
 */
export function isB2BCustomerType(customerType?: string | null): boolean {
  return customerType === 'dealer' || customerType === 'academy';
}

export function isB2CCustomerType(customerType?: string | null): boolean {
  return !isB2BCustomerType(customerType);
}
