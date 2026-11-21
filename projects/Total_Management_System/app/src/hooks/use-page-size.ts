'use client';

import { useEffect, useRef, useState } from 'react';

/**
 * 한 쪽에 보여 줄 건수를 화면 높이에 맞춘다 (2026-11-21)
 *
 * 세로로 긴 화면(세워 둔 모니터 등)에서 20건만 보이고 아래가 비던 것을 채운다.
 *  - 보통 화면(높이 1100 안팎까지)과 휴대폰은 지금처럼 20건 — 달라지는 것이 없다.
 *  - 줄 높이 약 42px · 목록 위아래 자리 약 290px 로 어림해 5건 단위로 올림하지 않고 내림(스크롤이 생기지 않게).
 *  - 20건보다 줄이지 않는다 · 60건을 넘기지 않는다(한 번에 받는 양).
 *  - 창 크기를 바꾸는 동안에는 기다렸다가 한 번만 바꾼다(그때마다 다시 조회하지 않게).
 */
export const PAGE_SIZE_BASE = 20;
const ROW_PX = 42;
const CHROME_PX = 290;
const MAX = 60;

export function pageSizeFor(innerHeight: number, innerWidth: number): number {
  if (innerWidth < 1024) return PAGE_SIZE_BASE; // 휴대폰 · 좁은 화면은 그대로
  const rows = Math.floor((innerHeight - CHROME_PX) / ROW_PX / 5) * 5;
  return Math.max(PAGE_SIZE_BASE, Math.min(MAX, rows));
}

/** onChange: 건수가 바뀌었을 때(보통 첫 쪽으로 돌리는 데 쓴다) */
export function usePageSize(onChange?: () => void): number {
  const [size, setSize] = useState(PAGE_SIZE_BASE);
  const current = useRef(PAGE_SIZE_BASE);
  const onChangeRef = useRef(onChange);
  useEffect(() => { onChangeRef.current = onChange; });
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | null = null;
    const apply = () => {
      const next = pageSizeFor(window.innerHeight, window.innerWidth);
      if (next === current.current) return;
      current.current = next;
      setSize(next);
      onChangeRef.current?.();
    };
    const onResize = () => { if (timer) clearTimeout(timer); timer = setTimeout(apply, 300); };
    apply();
    window.addEventListener('resize', onResize);
    return () => { window.removeEventListener('resize', onResize); if (timer) clearTimeout(timer); };
  }, []);
  return size;
}
