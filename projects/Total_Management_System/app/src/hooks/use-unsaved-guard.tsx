'use client';

import { useEffect, useRef, useState, type ReactNode } from 'react';
import { useRouter } from 'next/navigation';
import { DiscardAsk } from '@/components/ui/discard-ask';

/**
 * 저장하지 않은 내용 지키기 (2026-11-15) — 판매 입력 · 발주 작성 · 계약서
 *
 * 넣다가 다른 메뉴를 누르면 "저장하지 않은 내용이 있습니다" 를 한 번 묻는다.
 *   · 화면 안의 링크(사이드 메뉴 · 아래 메뉴 · 목록으로)를 눌렀을 때 → 묻는 창
 *   · 새로고침 · 탭 닫기 → 브라우저의 기본 경고
 * 저장에 성공해 화면이 스스로 이동하는 경우(router.push)는 묻지 않는다.
 * 브라우저의 뒤로 가기 버튼은 여기서 막지 못한다.
 *
 * 사용:  const guard = useUnsavedGuard(cart.length > 0 || !!memo.trim());   …   {guard}
 */
export function useUnsavedGuard(dirty: boolean): ReactNode {
  const router = useRouter();
  const dirtyRef = useRef(dirty);
  const [pending, setPending] = useState<string | null>(null);
  useEffect(() => { dirtyRef.current = dirty; }, [dirty]);

  useEffect(() => {
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (!dirtyRef.current) return;
      e.preventDefault();
      e.returnValue = '';
    };
    const onClick = (e: MouseEvent) => {
      if (!dirtyRef.current || e.defaultPrevented) return;
      if (e.button !== 0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;   // 새 탭 열기는 내용이 남으므로 그대로
      const a = (e.target as Element | null)?.closest?.('a[href]') as HTMLAnchorElement | null;
      if (!a || a.target === '_blank' || a.hasAttribute('download')) return;
      let url: URL;
      try { url = new URL(a.href, window.location.href); } catch { return; }
      if (url.origin !== window.location.origin) return;
      if (url.pathname === window.location.pathname) return;   // 같은 화면 안 이동
      e.preventDefault();
      e.stopPropagation();
      setPending(url.pathname + url.search + url.hash);
    };
    window.addEventListener('beforeunload', onBeforeUnload);
    document.addEventListener('click', onClick, true);   // 화면 부품보다 먼저 받아야 이동을 멈출 수 있다
    return () => {
      window.removeEventListener('beforeunload', onBeforeUnload);
      document.removeEventListener('click', onClick, true);
    };
  }, []);

  // 묻는 창이 떠 있는 동안 Esc = 계속 작성
  useEffect(() => {
    if (!pending) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') { e.stopPropagation(); setPending(null); } };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [pending]);

  if (!pending) return null;
  return (
    <DiscardAsk
      fixed
      title="저장하지 않은 내용이 있습니다"
      desc="다른 화면으로 가면 지금까지 넣은 내용이 사라집니다."
      keepLabel="계속 작성"
      discardLabel="버리고 이동"
      onKeep={() => setPending(null)}
      onDiscard={() => { const to = pending; dirtyRef.current = false; setPending(null); router.push(to); }}
    />
  );
}
