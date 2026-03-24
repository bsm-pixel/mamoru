'use client';

import { useQuery } from '@tanstack/react-query';

export interface LotteTrackResult {
  ok: boolean;
  state: string;
  raw: {
    tracking?: Array<{
      godsStatCd?: string;
      statTm?: string;
      orgNm?: string;
      statDt?: string;
      [key: string]: unknown;
    }>;
    [key: string]: unknown;
  };
}

/** 롯데택배 ALPS 배송 추적 */
export function useDeliveryTrack(invNo: string | null) {
  return useQuery<LotteTrackResult>({
    queryKey: ['delivery-track', invNo],
    queryFn: async () => {
      const res = await fetch(`/api/lotte/track?invNo=${invNo}`);
      if (!res.ok) throw new Error('추적 실패');
      return res.json();
    },
    enabled: !!invNo,
    staleTime: 60_000,
    refetchInterval: 300_000,
  });
}
