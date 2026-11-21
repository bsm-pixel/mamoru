'use client';

import { useEffect, useState } from 'react';
import { usePathname } from 'next/navigation';
import { UnifiedPrepModal } from '@/components/lazy-modals';

/**
 * 전체 출고 준비표를 여는 한 곳 (2026-11-16)
 *
 * 전엔 같은 버튼이 주문관리 · 판매 관리 · 복원수리 세 화면에 따로 있었다(창은 하나인데 문이 셋).
 * 이제 버튼은 머리줄(Topbar)에 하나만 있고, 어느 화면에서나 열린다. 창은 대시보드 레이아웃에 한 번만 둔다.
 * 처음 보이는 탭은 지금 있는 화면을 따른다(복원수리 화면 → 복원수리 탭).
 */
const OPEN_EVENT = 'mm:unified-prep';
export function openUnifiedPrep() {
  if (typeof window !== 'undefined') window.dispatchEvent(new CustomEvent(OPEN_EVENT));
}

export function UnifiedPrepHost() {
  const [open, setOpen] = useState(false);
  const pathname = usePathname();
  useEffect(() => {
    const onOpen = () => setOpen(true);
    window.addEventListener(OPEN_EVENT, onOpen);
    return () => window.removeEventListener(OPEN_EVENT, onOpen);
  }, []);
  if (!open) return null;
  const tab = pathname.startsWith('/repairs') ? 'repair' : pathname.startsWith('/sales') ? 'sale' : 'order';
  return <UnifiedPrepModal initialTab={tab} onClose={() => setOpen(false)} />;
}
