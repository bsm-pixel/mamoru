'use client';

import { useEffect, useRef, useState } from 'react';
import { attachModalKeys, type ModalKeys } from '@/lib/ui/modal-keys';
import { DiscardAsk } from './discard-ask';

/**
 * ESC 로 모달 닫기 (2026-09-16) + 입력 창 공통 키보드 동작 (2026-11-15)
 *
 * 왜 컴포넌트인가: 모달 배경 37곳이 제각각 손으로 짜여 있어, 각 컴포넌트 본문에 훅을 끼워 넣는 것보다
 * **배경 안에 이 한 줄을 넣는 편**이 안전하고 빠뜨릴 여지가 없다. 모달이 열릴 때만 마운트된다.
 *
 * 🔑 중첩 모달: 스택의 **맨 위(가장 나중에 열린 것)만** 닫는다.
 *    안 그러면 모달 위의 확인창에서 ESC 를 눌렀을 때 뒤의 모달까지 한꺼번에 닫힌다.
 *
 * 2026-11-15 이 한 줄이 같이 해 주는 것 (lib/ui/modal-keys.ts)
 *   · 열리면 커서가 첫 입력 칸으로, 닫히면 연 버튼으로, Tab 은 창 안에서만
 *   · Ctrl+Enter = 저장, 한 줄 칸에서 Enter = 다음 칸
 *   · 입력한 내용이 있으면 Esc 때 "입력한 내용을 버릴까요?" 한 번 묻기
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
  const anchor = useRef<HTMLSpanElement>(null);
  const keys = useRef<ModalKeys | null>(null);
  const [asking, setAsking] = useState(false);
  const askingRef = useRef(false);
  const onCloseRef = useRef(onClose);
  useEffect(() => { onCloseRef.current = onClose; }, [onClose]);
  useEffect(() => { askingRef.current = asking; }, [asking]);

  useEffect(() => {
    if (disabled) return;
    ensureBinding();
    const fn = () => {
      if (askingRef.current) { setAsking(false); return; }          // 묻는 중의 Esc = 계속 입력
      if (keys.current?.isDirty()) { setAsking(true); return; }     // 입력한 내용이 있으면 한 번 묻는다
      onCloseRef.current();
    };
    stack.push(fn);
    const root = anchor.current?.parentElement || null;
    if (root) keys.current = attachModalKeys(root, { trapTab: true, isTop: () => stack[stack.length - 1] === fn });
    return () => {
      const i = stack.lastIndexOf(fn);
      if (i >= 0) stack.splice(i, 1);
      keys.current?.dispose();
      keys.current = null;
    };
  }, [disabled]);

  return (
    <>
      <span ref={anchor} hidden data-modal-ignore />
      {asking && <DiscardAsk fixed onKeep={() => setAsking(false)} onDiscard={() => { setAsking(false); onCloseRef.current(); }} />}
    </>
  );
}
