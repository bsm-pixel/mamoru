'use client';

import { useEffect, useRef } from 'react';

/**
 * "입력한 내용을 버릴까요?" — Esc 로 창을 닫을 때 입력한 내용이 있으면 한 번 묻는다 (2026-11-15)
 * 커서는 안전한 쪽([계속 입력])에 둔다. 이 창에서 Esc = 계속 입력.
 * 창 본체의 Tab·Enter 처리에 끼지 않도록 data-modal-ignore 를 단다.
 */
export function DiscardAsk({ onKeep, onDiscard, fixed, title = '입력한 내용을 버릴까요?', desc = '닫으면 지금까지 넣은 내용이 사라집니다.', keepLabel = '계속 입력', discardLabel = '버리고 닫기' }: {
  onKeep: () => void; onDiscard: () => void; fixed?: boolean;
  title?: string; desc?: string; keepLabel?: string; discardLabel?: string;
}) {
  const keepRef = useRef<HTMLButtonElement>(null);
  useEffect(() => { keepRef.current?.focus(); }, []);
  return (
    <div
      data-modal-ignore
      data-modal-open=""
      className={`${fixed ? 'fixed' : 'absolute'} inset-0 z-[200] flex items-center justify-center bg-black/40 p-4`}
      onMouseDown={(e) => e.stopPropagation()}
      onMouseUp={(e) => e.stopPropagation()}
      onClick={(e) => e.stopPropagation()}
      onKeyDown={(e) => {
        if (e.key === 'Tab') {
          // 두 버튼 사이에서만 돈다
          const btns = Array.from(e.currentTarget.querySelectorAll('button'));
          const i = btns.indexOf(document.activeElement as HTMLButtonElement);
          e.preventDefault();
          btns[(i + (e.shiftKey ? btns.length - 1 : 1)) % btns.length]?.focus();
        }
      }}
    >
      <div role="alertdialog" aria-label={title} className="w-full max-w-[320px] rounded-xl bg-white border border-neutral-200 p-4 text-left">
        <p className="text-sm font-bold text-neutral-900">{title}</p>
        <p className="mt-1 text-xs text-neutral-500">{desc}</p>
        <div className="mt-4 flex justify-end gap-2">
          <button ref={keepRef} type="button" data-safe onClick={onKeep}
            className="h-8 px-3 rounded-md text-sm font-semibold bg-stone-900 text-white hover:bg-stone-800">
            {keepLabel}
          </button>
          <button type="button" onClick={onDiscard}
            className="h-8 px-3 rounded-md text-sm font-semibold border border-neutral-200 text-neutral-600 hover:bg-neutral-50">
            {discardLabel}
          </button>
        </div>
      </div>
    </div>
  );
}
