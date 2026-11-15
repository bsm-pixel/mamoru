'use client';

import { useCallback, useSyncExternalStore } from 'react';

/**
 * 기억하는 상태 (2026-11-15) — 탭 · 기간처럼 "다시 들어와도 그대로였으면 하는 것"
 *
 * useState 자리에 그대로 쓴다. 값은 이 브라우저에 저장된다(다른 기기와는 따로).
 *   const [tab, setTab] = useStickyState<SalesTab>('sales.tab', 'all');
 *
 * allowed 를 주면 저장된 값이 그 안에 없을 때 기본값으로 돌아간다(탭 이름이 바뀌어도 화면이 깨지지 않게).
 * 처음 그릴 때는 기본값으로 그리고 곧바로 저장된 값으로 바뀐다(서버가 그린 화면과 어긋나지 않게).
 */
const EVT = 'mm:sticky-state';
const PREFIX = 'tms.sticky.';

function subscribe(cb: () => void) {
  window.addEventListener(EVT, cb);
  window.addEventListener('storage', cb);
  return () => { window.removeEventListener(EVT, cb); window.removeEventListener('storage', cb); };
}

// 저장소를 못 쓰는 환경(사생활 보호 창 등)에서도 화면 안에서는 값이 바뀌도록 메모리에도 둔다
const mem = new Map<string, string>();

function read(key: string): string | null {
  if (mem.has(key)) return mem.get(key)!;
  try { return window.localStorage.getItem(PREFIX + key); } catch { return null; }
}

export function useStickyState<T extends string>(key: string, initial: T, allowed?: readonly T[]): [T, (v: T) => void] {
  const raw = useSyncExternalStore(subscribe, () => read(key), () => null);
  const value = raw !== null && (!allowed || (allowed as readonly string[]).includes(raw)) ? (raw as T) : initial;

  const set = useCallback((v: T) => {
    mem.set(key, v);
    try { window.localStorage.setItem(PREFIX + key, v); } catch { /* 저장 못 해도 화면은 그대로 쓴다 */ }
    window.dispatchEvent(new Event(EVT));
  }, [key]);

  return [value, set];
}
