'use client';

import { useQuery } from '@tanstack/react-query';
import type { TrackRecord } from '@/lib/lotte/track-steps';

export interface LotteTrackResult {
  ok: boolean;
  state: string;
  raw: {
    // 필드명은 실측 기준(2026-09-26) — 옛 타입의 statDt/statTm/orgNm 은 응답에 없는 이름이었다
    tracking?: TrackRecord[];
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
