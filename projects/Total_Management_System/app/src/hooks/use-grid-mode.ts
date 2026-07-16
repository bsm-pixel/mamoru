'use client';

import { useState, useEffect } from 'react';
// (useGridMode 유지: 향후 토글 필요 화면용 / useIsLg: 그리드 전용 화면용)

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

/**
 * isLg(≥1024px) 감지만 — 그리드 전용 화면(토글 없음)에서 PC/모바일 레이아웃 분기용.
 * (2026-07-16 카드보기·토글 폐지 후 마스터-디테일 페이지들이 사용)
 */
export function useIsLg() {
  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1024px)');
    setIsLg(mq.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, []);
  return isLg;
}
