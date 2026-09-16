'use client';

import { useEffect } from 'react';

/**
 * ESC 로 모달 닫기 (2026-09-16)
 *
 * 왜 컴포넌트인가: 모달 배경 37곳이 제각각 손으로 짜여 있어, 각 컴포넌트 본문에 훅을 끼워 넣는 것보다
 * **배경 안에 이 한 줄을 넣는 편**이 안전하고 빠뜨릴 여지가 없다. 모달이 열릴 때만 마운트된다.
 *
 * 🔑 중첩 모달: 스택의 **맨 위(가장 나중에 열린 것)만** 닫는다.
 *    안 그러면 모달 위의 확인창에서 ESC 를 눌렀을 때 뒤의 모달까지 한꺼번에 닫힌다.
 *
 * 사용:
 *   <div className="fixed inset-0 …" {...backdropClose(onClose)}>
 *     <EscClose onClose={onClose} />
 *     …
 *   </div>
 *
 * ⚠️ 작성 중 내용을 지켜야 하는 모달(검수 등)은 쓰지 않는다 — 공용 `ui/modal.tsx` 의 preventAutoClose 사용.
 */

// 열려 있는 ESC 대상들의 스택 (맨 뒤 = 가장 위에 뜬 모달)
const stack: Array<() => void> = [];

let bound = false;
function ensureBinding() {
  if (bound || typeof window === 'undefined') return;
  bound = true;
  window.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape' || stack.length === 0) return;
    // IME 조합 중(한글 입력)에는 무시 — 조합 취소용 ESC 가 모달을 닫으면 안 된다
    if ((e as KeyboardEvent).isComposing) return;
    e.stopPropagation();
    stack[stack.length - 1]();
  });
}

export function EscClose({ onClose, disabled }: { onClose: () => void; disabled?: boolean }) {
  useEffect(() => {
    if (disabled) return;
    ensureBinding();
    const fn = () => onClose();
    stack.push(fn);
    return () => {
      const i = stack.lastIndexOf(fn);
      if (i >= 0) stack.splice(i, 1);
    };
  }, [onClose, disabled]);

  return null;
}
