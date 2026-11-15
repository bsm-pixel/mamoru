'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import { todayKST, addDays } from '@/lib/utils/kst';

/**
 * 자주 파는 제품 (2026-11-15) — 최근 90일 판매 수량이 많은 순서의 제품 id
 * 판매 입력에서 제품 전체(160여 개)를 깔아 두는 대신 이 10개만 먼저 보여 준다.
 * 조회에 실패해도 화면은 그대로 쓸 수 있게 빈 목록을 돌려준다.
 */
export function useTopProductIds(days = 90, limit = 10) {
  return useQuery({
    queryKey: ['top-products', days, limit],
    staleTime: 10 * 60_000,
    queryFn: async (): Promise<string[]> => {
      try {
        const from = addDays(todayKST(), -days);
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data, error } = await (createClient() as any)
          .from('offline_sale_items')
          .select('product_id, quantity, offline_sales!inner(sale_date, cancelled_at)')
          .gte('offline_sales.sale_date', from)
          .is('offline_sales.cancelled_at', null)
          .not('product_id', 'is', null)
          .limit(3000);
        if (error) return [];
        const sum: Record<string, number> = {};
        for (const r of (data || []) as Array<{ product_id: string; quantity: number | null }>) {
          sum[r.product_id] = (sum[r.product_id] || 0) + (r.quantity || 1);
        }
        return Object.entries(sum).sort((a, b) => b[1] - a[1]).slice(0, limit).map(([id]) => id);
      } catch {
        return [];
      }
    },
  });
}
