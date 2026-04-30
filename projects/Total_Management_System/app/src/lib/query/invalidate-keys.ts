import type { QueryClient } from '@tanstack/react-query';

/**
 * 매출/돈 흐름에 영향을 주는 모든 query key 모음.
 * sale/repair/expense/delivery mutation의 onSuccess에서 invalidateFinancialQueries(qc) 한 줄로 일괄 무효화.
 *
 * 075 — invalidation 누락으로 대시보드가 60s 동안 옛 수치를 보여주던 문제 해결.
 * 새 query 추가 시 이 배열에 동참시켜야 cross-domain 갱신이 자동으로 반영됨.
 */
export const FINANCIAL_QUERY_KEYS = [
  'sales',
  'sales-stats',
  'sales-tab-counts',
  'hub-stats',
  'order-dashboard-stats',
  'repair-dashboard-stats',
  'products',
  'inventory',
  'customers',
  'expenses',
  'recurring-expenses',
  'deliveries',
  'repairs',
  'repair-tabs',
  'orders',
] as const;

export function invalidateFinancialQueries(queryClient: QueryClient) {
  for (const key of FINANCIAL_QUERY_KEYS) {
    queryClient.invalidateQueries({ queryKey: [key] });
  }
}
