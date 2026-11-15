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
 *
 * ⚠️ 모바일의 옆에서 밀려 나오는 창(ui/slide-panel.tsx)은 열릴 때 자기 기록을 한 칸 쌓고, 닫힐 때 그 칸을 history.back() 으로 치운다.
 *    그 back() 이 "?id= 가 붙은 칸"으로 돌아오므로, 닫는 순간 주소만 고치면 방금 닫은 건이 다시 열린다.
 *    → 그 경우에는 먼저 "닫힘"으로 보이게 하고(pendingClear), back() 이 끝난 뒤에 주소에서 뺀다.
 */
const EVT = 'mm:url-selection';
const pendingClear = new Set<string>();

function notify() {
  window.dispatchEvent(new Event(EVT));
}

function stripParam(key: string) {
  const url = new URL(window.location.href);
  if (!url.searchParams.has(key)) return;
  url.searchParams.delete(key);
  window.history.replaceState(window.history.state, '', url);
}

let popBound = false;
function bindPop() {
  if (popBound) return;
  popBound = true;
  window.addEventListener('popstate', () => {
    if (pendingClear.size === 0) return;
    // 밀려 나온 창이 자기 칸을 치우고 돌아온 직후 — 이제 주소에서 뺀다
    for (const key of Array.from(pendingClear)) stripParam(key);
    pendingClear.clear();
    notify();
  });
}

function subscribe(cb: () => void) {
  bindPop();
  window.addEventListener('popstate', cb);
  window.addEventListener(EVT, cb);
  return () => { window.removeEventListener('popstate', cb); window.removeEventListener(EVT, cb); };
}

export function useUrlSelection(key = 'id'): [string | null, (id: string | null) => void] {
  const value = useSyncExternalStore(
    subscribe,
    () => (pendingClear.has(key) ? null : new URLSearchParams(window.location.search).get(key)),
    () => null,
  );

  const set = useCallback((id: string | null) => {
    const url = new URL(window.location.href);
    const cur = pendingClear.has(key) ? null : url.searchParams.get(key);
    if (id === cur) return;
    if (id) {
      pendingClear.delete(key);
      url.searchParams.set(key, id);
      // 기록은 한 번만 쌓는다 — 줄을 옮겨 다니거나 닫았다 다시 열 때마다 쌓이면 뒤로 가기를 여러 번 눌러야 한다
      if (window.history.state?.mmSel) window.history.replaceState(window.history.state, '', url);
      else window.history.pushState({ ...window.history.state, mmSel: true }, '', url);
    } else if ((window.history.state as { __slidePanel?: boolean } | null)?.__slidePanel) {
      // 밀려 나온 창이 곧 back() 한다 → 지금은 닫힌 것으로만 보이게 하고, 돌아온 뒤 주소에서 뺀다
      pendingClear.add(key);
      setTimeout(() => {
        // back() 이 오지 않은 경우의 안전망
        if (pendingClear.has(key)) { pendingClear.delete(key); stripParam(key); notify(); }
      }, 600);
    } else {
      // 닫기: 주소에서 바로 뺀다(닫자마자 다른 줄을 눌러도 꼬이지 않는다)
      stripParam(key);
    }
    notify();
  }, [key]);

  return [value, set];
}
