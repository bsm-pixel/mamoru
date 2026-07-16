'use client';

import { useState, useEffect } from 'react';

/**
 * PC 그리드(밀집 표) 토글 + isLg(≥1024px) 감지 공용 훅 (2026-07-16)
 *
 * 판매관리에서 검증된 패턴을 복원수리·고객·매입 등으로 일반화.
 * 페이지별 선호는 storageKey(localStorage)로 각각 저장한다.
 *   예) 'sales-pc-grid' · 'repairs-pc-grid' · 'customers-pc-grid' · 'purchasing-pc-grid'
 *
 * SSR 안전: 초기값 false → mount effect에서 matchMedia·localStorage 읽음.
 */
export function useGridMode(storageKey: string) {
  const [isLg, setIsLg] = useState(false);
  const [gridMode, setGridMode] = useState(false);

  const toggleGrid = () =>
    setGridMode((v) => {
      const n = !v;
      try {
        localStorage.setItem(storageKey, n ? '1' : '0');
      } catch {
        /* noop */
      }
      return n;
    });

  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1024px)');
    setIsLg(mq.matches);
    try {
      setGridMode(localStorage.getItem(storageKey) === '1');
    } catch {
      /* noop */
    }
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, [storageKey]);

  return { isLg, gridMode, toggleGrid };
}
