'use client';

import { useCallback } from 'react';
import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';

/** 한 고객의 활동유형 (아임웹주문 / 오프라인구매 / 복원수리) */
export interface ActivityTypes {
  has_order: boolean;
  has_sale: boolean;
  has_repair: boolean;
}

function normalizePhone(phone?: string | null): string {
  return (phone || '').replace(/\D/g, '');
}

/**
 * 여러 전화번호의 활동유형을 배치로 조회 (RPC get_activity_types_by_phones).
 * 목록 화면에서 보이는 고객들의 phone 을 모아 한 번에 조회 → 행마다 칩 렌더.
 * @returns (phone) => ActivityTypes | undefined  — 정규화 전화번호로 조회하는 lookup 함수
 */
export function useActivityTypes(phones: (string | null | undefined)[]) {
  const supabase = createClient();
  const normalized = Array.from(new Set(phones.map(normalizePhone).filter(Boolean)));
  const key = normalized.slice().sort().join(',');

  const { data } = useQuery({
    queryKey: ['activity-types', key],
    enabled: normalized.length > 0,
    staleTime: 60_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any).rpc('get_activity_types_by_phones', { phones: normalized });
      if (error) throw error;
      const map: Record<string, ActivityTypes> = {};
      for (const row of (data || []) as Array<ActivityTypes & { phone: string }>) {
        map[row.phone] = { has_order: row.has_order, has_sale: row.has_sale, has_repair: row.has_repair };
      }
      return map;
    },
  });

  return useCallback(
    (phone?: string | null): ActivityTypes | undefined => {
      if (!data) return undefined;
      return data[normalizePhone(phone)];
    },
    [data]
  );
}
