'use client';

import { useState, useEffect, useCallback } from 'react';

/** PC 여부 감지 (lg: 1024px+) — SSR-safe */
export function useIsLg() {
  const [isLg, setIsLg] = useState(false);

  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  return isLg;
}

/** ESC 키로 콜백 실행 (상세 패널 닫기 등) */
export function useEscapeKey(callback: () => void, enabled = true) {
  const handler = useCallback((e: KeyboardEvent) => {
    if (e.key === 'Escape') callback();
  }, [callback]);

  useEffect(() => {
    if (!enabled) return;
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [handler, enabled]);
}
