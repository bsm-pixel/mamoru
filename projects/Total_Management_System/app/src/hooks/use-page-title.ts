'use client';

import { useEffect } from 'react';

/**
 * 브라우저 탭 이름 (2026-11-15) — 전엔 모든 화면이 「MAMORU TMS」라 탭을 여러 개 열면 구분이 안 됐다.
 *   「판매 관리 · TMS」, 처리할 건이 있으면 「판매 관리 (3) · TMS」
 * 머리줄(Topbar)이 이 훅을 부르므로 화면은 Topbar 의 title · count 만 주면 된다.
 * 머리줄이 없는 화면만 직접 부른다.
 *
 * 2026-11-16 건수가 없는 화면(고객 · 제품 · 설정 등)에서 이름이 다시 「MAMORU TMS」로 돌아가던 문제:
 *   주소로 바로 열거나 새로 고치면, 화면이 뜬 직후 Next 가 기본 제목(layout 의 metadata)을 한 번 더 써서 우리 이름을 덮었다.
 *   건수가 있는 화면은 건수가 도착할 때 다시 써서 티가 안 났을 뿐이다.
 *   → <title> 이 바뀌는지 지켜보다가, 우리 이름이 아니면 다시 쓴다. (화면 점검 npm run check:screen 이 17개 화면을 확인한다)
 */
let wanted: string | null = null;
let watching = false;

function apply() {
  if (wanted && document.title !== wanted) document.title = wanted;
}

function watch() {
  if (watching || typeof MutationObserver === 'undefined') return;
  watching = true;
  new MutationObserver(apply).observe(document.head, { childList: true, subtree: true, characterData: true });
}

export function usePageTitle(title: string, count?: number | null) {
  useEffect(() => {
    if (!title) return;
    const n = typeof count === 'number' && count > 0 ? ` (${count > 999 ? '999+' : count})` : '';
    const mine = `${title}${n} · TMS`;
    wanted = mine;
    watch();
    apply();
    // 떠날 때 내 이름을 내려놓는다 — 머리줄이 없는 화면(로그인 등)까지 옛 이름이 따라가지 않게
    return () => { if (wanted === mine) wanted = null; };
  }, [title, count]);
}
