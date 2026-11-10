'use client';

import { useEffect, useState } from 'react';

/**
 * 현재 시각(ms)을 주기적으로 갱신 — 렌더 중 Date.now() 직접 호출 금지(react-hooks/purity) 대체용.
 * 처음 렌더는 0 → 마운트 직후 실제 시각으로 채운다.
 */
export function useNow(intervalMs = 5000): number {
  const [now, setNow] = useState(0);
  useEffect(() => {
    const first = setTimeout(() => setNow(Date.now()), 0);
    const id = setInterval(() => setNow(Date.now()), intervalMs);
    return () => { clearTimeout(first); clearInterval(id); };
  }, [intervalMs]);
  return now;
}
