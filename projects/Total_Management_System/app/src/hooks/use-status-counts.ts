'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';

/**
 * 상태 탭 건수 (2026-11-15) — 반품·교환 · 매입관리 탭이 건수 없이 「–」로만 보였다
 * 표 하나의 status 칸만 가볍게 읽어 상태별로 센다. 전체 건수는 total.
 *   const { counts, total } = useStatusCounts('purchase_orders');   counts['ordered']
 */
export function useStatusCounts(table: 'purchase_orders' | 'returns') {
  const { data } = useQuery({
    queryKey: ['status-counts', table],
    staleTime: 30_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (createClient() as any).from(table).select('status').limit(5000);
      if (error) throw error;
      const counts: Record<string, number> = {};
      for (const r of (data || []) as Array<{ status: string | null }>) {
        const k = r.status || '';
        counts[k] = (counts[k] || 0) + 1;
      }
      return { counts, total: (data || []).length };
    },
  });
  return { counts: data?.counts || {}, total: data?.total ?? null, loaded: !!data };
}
