'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import { formatDateTimeShort } from '@/lib/utils/format';

/**
 * 임시 저장 (2026-11-21) — 판매 입력 · 발주 작성 · 계약서 작성
 *
 * 쓰던 내용을 이 기기의 브라우저에 남겨, 창을 닫거나 새로고침해도 이어서 쓸 수 있게 한다.
 *  - 다시 열면 바로 채우지 않고 "작성하던 내용이 있습니다 [이어서 쓰기] [버리기]"를 먼저 묻는다.
 *    (묻기 전에는 저장하지 않는다 — 빈 화면이 남은 초안을 덮어쓰지 않게)
 *  - 하루(24시간)가 지난 초안 · 형식이 바뀐 초안은 버린다. 재고 · 가격은 그때 기준이라 오래되면 어긋난다.
 *  - 저장에 성공하면 clear() 로 지운다(같은 내용이 두 번 들어가지 않게).
 *  - 서명 · 손글씨 이미지는 넣지 않는다(크고, 함께 쓰는 태블릿에 남으면 안 된다). 무엇을 넣을지는 화면이 고른다.
 */
const MAX_AGE_MS = 24 * 60 * 60 * 1000;

interface Stored<T> { v: number; at: number; data: T }

export function useFormDraft<T>(key: string, version: number, snapshot: T, hasContent: boolean, enabled = true) {
  const storageKey = `tms.draft.${key}`;
  const [pending, setPending] = useState<Stored<T> | null>(null);
  const [ready, setReady] = useState(false);
  const stopped = useRef(false);

  // 처음 한 번: 남은 초안이 있는지 본다
  useEffect(() => {
    const id = requestAnimationFrame(() => {
      if (enabled) {
        try {
          const raw = localStorage.getItem(storageKey);
          if (raw) {
            const d = JSON.parse(raw) as Stored<T>;
            if (d && d.v === version && typeof d.at === 'number' && Date.now() - d.at < MAX_AGE_MS) setPending(d);
            else localStorage.removeItem(storageKey);
          }
        } catch { /* 저장소를 못 쓰는 브라우저 — 임시 저장 없이 그대로 쓴다 */ }
      }
      setReady(true);
    });
    return () => cancelAnimationFrame(id);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 바뀔 때마다(잠깐 기다렸다가) 저장
  const json = JSON.stringify(snapshot);
  useEffect(() => {
    if (!ready || !enabled || pending || stopped.current) return;
    const t = setTimeout(() => {
      if (stopped.current) return;
      try {
        if (hasContent) localStorage.setItem(storageKey, JSON.stringify({ v: version, at: Date.now(), data: JSON.parse(json) } satisfies Stored<unknown>));
        else localStorage.removeItem(storageKey);
      } catch { /* 용량 초과 등 — 무시 */ }
    }, 500);
    return () => clearTimeout(t);
  }, [json, hasContent, ready, enabled, pending, storageKey, version]);

  /** [이어서 쓰기] — 초안을 돌려준다(화면이 각 칸에 채운다) */
  const take = useCallback((): T | null => {
    const d = pending?.data ?? null;
    setPending(null);
    return d;
  }, [pending]);
  /** [버리기] */
  const discard = useCallback(() => {
    try { localStorage.removeItem(storageKey); } catch { /* 무시 */ }
    setPending(null);
  }, [storageKey]);
  /** 저장 성공 뒤 — 초안을 지우고 더 저장하지 않는다 */
  const clear = useCallback(() => {
    stopped.current = true;
    try { localStorage.removeItem(storageKey); } catch { /* 무시 */ }
  }, [storageKey]);

  return { pendingAt: pending?.at ?? null, take, discard, clear };
}

/** "작성하던 내용이 있습니다" 안내 줄 */
export function DraftBanner({ at, onResume, onDiscard }: { at: number | null; onResume: () => void; onDiscard: () => void }) {
  if (!at) return null;
  return (
    <div data-draft-banner className="flex flex-wrap items-center gap-2 rounded-lg border border-neutral-300 bg-white px-3 py-2 text-xs">
      <span className="font-semibold text-stone-900">작성하던 내용이 있습니다</span>
      <span className="text-neutral-500">{formatDateTimeShort(new Date(at).toISOString())} · 재고와 가격은 그때 기준입니다</span>
      <span className="ml-auto flex items-center gap-1.5">
        <button type="button" onClick={onDiscard} className="rounded-md px-2 py-1 font-semibold text-neutral-500 hover:text-neutral-800">버리기</button>
        <button type="button" onClick={onResume} className="rounded-md bg-stone-900 px-2.5 py-1 font-semibold text-white hover:bg-stone-800">이어서 쓰기</button>
      </span>
    </div>
  );
}
