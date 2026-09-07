'use client';

import { useEffect, useRef } from 'react';

export interface Hotkey {
  /** 소문자 조합키: 'f2', 'f4', 'ctrl+s', 'escape', '/', '?' */
  combo: string;
  handler: (e: KeyboardEvent) => void;
  /** input/textarea/select/contenteditable 포커스 중에도 동작할지 (기본 false) */
  allowInInput?: boolean;
  /** 브라우저 기본 동작 취소 (기본 true) */
  preventDefault?: boolean;
}

/** 타이핑 대상(입력요소)에 포커스 중인지 */
function isTyping(t: EventTarget | null): boolean {
  const el = t as HTMLElement | null;
  if (!el || !el.tagName) return false;
  const tag = el.tagName;
  return tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT' || el.isContentEditable;
}

/** KeyboardEvent → 소문자 조합키. 단일 문자키는 shift 를 키 자체(예: '?')가 반영하므로 제외 */
function comboOf(e: KeyboardEvent): string {
  const k = e.key.toLowerCase();
  const parts: string[] = [];
  if (e.ctrlKey || e.metaKey) parts.push('ctrl');
  if (e.altKey) parts.push('alt');
  if (e.shiftKey && k.length > 1) parts.push('shift');
  parts.push(k);
  return parts.join('+');
}

/**
 * 전역 keydown 단축키 훅.
 * - 리스너는 컴포넌트당 1개만 등록, 바인딩은 ref 로 최신화 → 매 렌더 재구독 없음(버벅임/누수 방지)
 * - 페이지 컴포넌트에서 쓰면 그 화면이 떠 있을 때만 동작(언마운트 시 자동 해제)
 */
export function useHotkeys(hotkeys: Hotkey[], enabled = true) {
  const ref = useRef(hotkeys);
  // 렌더 중이 아니라 커밋 후 최신화 → 리스너 재구독 없이 항상 최신 핸들러 참조
  useEffect(() => { ref.current = hotkeys; });

  useEffect(() => {
    if (!enabled) return;
    const onKeyDown = (e: KeyboardEvent) => {
      if (e.repeat) return; // 키 홀드 반복 무시
      const combo = comboOf(e);
      const typing = isTyping(e.target);
      for (const hk of ref.current) {
        if (hk.combo !== combo) continue;
        if (typing && !hk.allowInInput) continue;
        if (hk.preventDefault !== false) e.preventDefault();
        hk.handler(e);
        return; // 첫 매칭만 실행
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [enabled]);
}

/** 모달(native <dialog>)이 열려 있으면 true — 페이지 Esc 단축키가 모달에 양보하도록 판정 */
export function isDialogOpen(): boolean {
  if (typeof document === 'undefined') return false;
  return !!document.querySelector('dialog[open]');
}
