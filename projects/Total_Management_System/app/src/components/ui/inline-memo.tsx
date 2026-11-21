'use client';

import { useState } from 'react';

/**
 * 목록에서 메모 바로 고치기 (2026-11-21)
 *
 * 줄에 한 줄로 보이다가 누르면 그 자리에서 입력칸이 된다. Enter = 저장 · Esc = 그만두기 · 다른 곳을 누르면 저장.
 *  - 줄을 누르는 동작(상세 열기)으로 번지지 않는다.
 *  - 메모가 없으면 줄에 마우스를 올렸을 때만 「+ 메모」가 보인다(DataGrid 줄의 group-hover).
 *  - 수량 · 금액처럼 다른 값이 함께 바뀌는 칸에는 쓰지 않는다(이력 · 재계산이 빠진다). 메모 전용.
 */
export function InlineMemo({ value, onSave, placeholder = '메모', className = '' }: {
  value: string | null | undefined;
  onSave: (next: string) => void | Promise<unknown>;
  placeholder?: string;
  className?: string;
}) {
  const current = value ?? '';
  const [draft, setDraft] = useState<string | null>(null); // null = 보기 상태
  const editing = draft !== null;

  const commit = () => {
    if (draft === null) return;
    const next = draft.trim();
    setDraft(null);
    if (next !== current.trim()) void onSave(next);
  };

  if (editing) {
    return (
      <input
        autoFocus
        data-nodirty
        data-inline-memo
        value={draft}
        maxLength={500}
        placeholder={placeholder}
        onChange={(e) => setDraft(e.target.value)}
        onClick={(e) => e.stopPropagation()}
        onBlur={commit}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === 'Enter' && !e.nativeEvent.isComposing) { e.preventDefault(); commit(); }
          if (e.key === 'Escape') { e.preventDefault(); setDraft(null); }
        }}
        className={`block w-full min-w-[120px] h-6 px-1.5 rounded border border-neutral-300 bg-white text-[11px] text-stone-800 focus:outline-none focus:ring-1 focus:ring-neutral-500 ${className}`}
      />
    );
  }
  return (
    <button
      type="button"
      data-inline-memo
      title={current ? `${current}\n(눌러서 고치기)` : '메모 적기'}
      onClick={(e) => { e.stopPropagation(); setDraft(current); }}
      className={`block max-w-full truncate text-left text-[11px] ${current ? 'text-neutral-500 hover:text-stone-900' : 'text-neutral-300 opacity-0 group-hover:opacity-100 focus:opacity-100 hover:text-neutral-600'} transition ${className}`}
    >
      {current ? `✎ ${current}` : `+ ${placeholder}`}
    </button>
  );
}
