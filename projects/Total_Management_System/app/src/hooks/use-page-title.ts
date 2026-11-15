'use client';

import { useEffect } from 'react';

/**
 * 브라우저 탭 이름 (2026-11-15) — 전엔 모든 화면이 「MAMORU TMS」라 탭을 여러 개 열면 구분이 안 됐다.
 *   「판매 관리 · TMS」, 처리할 건이 있으면 「판매 관리 (3) · TMS」
 * 머리줄(Topbar)이 이 훅을 부르므로 화면은 Topbar 의 title · count 만 주면 된다.
 * 머리줄이 없는 화면만 직접 부른다.
 */
export function usePageTitle(title: string, count?: number | null) {
  useEffect(() => {
    if (!title) return;
    const n = typeof count === 'number' && count > 0 ? ` (${count > 999 ? '999+' : count})` : '';
    document.title = `${title}${n} · TMS`;
  }, [title, count]);
}
