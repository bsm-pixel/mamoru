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

/**
 * 이 고객이 최근에 산 제품 (2026-11-27) — 판매 입력에서 고객을 먼저 고르면 그 고객의 재구매 제품을 목록 맨 위에 올린다.
 * 최근 판매(취소 제외) 20건 → 그 판매의 제품을 최근 순으로, 겹치지 않게 최대 limit 개. 실패하면 빈 목록(화면은 그대로).
 */
export function useCustomerRecentProductIds(customerId: string | null | undefined, limit = 5) {
  return useQuery({
    queryKey: ['customer-recent-products', customerId, limit],
    enabled: !!customerId,
    staleTime: 5 * 60_000,
    queryFn: async (): Promise<string[]> => {
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const db = createClient() as any;
        const { data: sales, error } = await db.from('offline_sales')
          .select('id, sale_date, created_at')
          .eq('customer_id', customerId)
          .is('cancelled_at', null)
          .order('sale_date', { ascending: false })
          .limit(20);
        if (error || !sales?.length) return [];
        const rows = sales as Array<{ id: string; sale_date: string | null; created_at: string | null }>;
        const rank = new Map(rows
          .slice().sort((a, b) => `${b.sale_date}${b.created_at}`.localeCompare(`${a.sale_date}${a.created_at}`))
          .map((s, i) => [s.id, i]));
        const { data: items, error: e2 } = await db.from('offline_sale_items')
          .select('product_id, sale_id')
          .in('sale_id', rows.map((s) => s.id))
          .not('product_id', 'is', null);
        if (e2) return [];
        const sorted = ((items || []) as Array<{ product_id: string | null; sale_id: string }>)
          .filter((r) => !!r.product_id)
          .sort((a, b) => (rank.get(a.sale_id) ?? 99) - (rank.get(b.sale_id) ?? 99));
        const out: string[] = [];
        for (const r of sorted) { if (!out.includes(r.product_id!)) out.push(r.product_id!); if (out.length >= limit) break; }
        return out;
      } catch {
        return [];
      }
    },
  });
}
