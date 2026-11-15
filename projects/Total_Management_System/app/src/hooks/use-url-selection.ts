'use client';

import { useCallback, useSyncExternalStore } from 'react';

/**
 * 고른 건을 주소에 붙이기 (2026-11-15)
 *
 * 전엔 상세를 열어도 주소가 그대로라 ① 그 건을 가리키는 링크를 만들 수 없고 ② 뒤로 가기를 누르면 화면을 떠났다.
 * 이제 상세를 열면 주소에 ?id=… 가 붙고, 뒤로 가기 = 상세 닫기, 그 주소로 들어오면 그 건이 열려 있다.
 *
 *   const [selectedId, setSelectedId] = useUrlSelection();     // useState<string | null>(null) 자리에 그대로
 *
 * 화면 전체를 다시 그리지 않도록 Next 라우터 대신 브라우저 history 를 직접 쓴다.
 */
const EVT = 'mm:url-selection';

function subscribe(cb: () => void) {
  window.addEventListener('popstate', cb);
  window.addEventListener(EVT, cb);
  return () => { window.removeEventListener('popstate', cb); window.removeEventListener(EVT, cb); };
}

export function useUrlSelection(key = 'id'): [string | null, (id: string | null) => void] {
  const value = useSyncExternalStore(
    subscribe,
    () => new URLSearchParams(window.location.search).get(key),
    () => null,
  );

  const set = useCallback((id: string | null) => {
    const url = new URL(window.location.href);
    const cur = url.searchParams.get(key);
    if (id === cur) return;
    if (id) {
      url.searchParams.set(key, id);
      // 기록은 한 번만 쌓는다 — 줄을 옮겨 다니거나 닫았다 다시 열 때마다 쌓이면 뒤로 가기를 여러 번 눌러야 한다
      if (window.history.state?.mmSel) window.history.replaceState(window.history.state, '', url);
      else window.history.pushState({ ...window.history.state, mmSel: true }, '', url);
    } else {
      // 닫기: 주소에서 바로 뺀다(뒤로 가기를 기다리지 않아, 닫자마자 다른 줄을 눌러도 꼬이지 않는다)
      url.searchParams.delete(key);
      window.history.replaceState(window.history.state, '', url);
    }
    window.dispatchEvent(new Event(EVT));
  }, [key]);

  return [value, set];
}
