'use client';

import { useEffect, useRef } from 'react';
import { X } from 'lucide-react';
import { cn } from '@/lib/utils/cn';

interface SlidePanelProps {
  open: boolean;
  onClose: () => void;
  children: React.ReactNode;
  title?: string;
  className?: string;
}

export function SlidePanel({ open, onClose, children, title, className }: SlidePanelProps) {
  // onClose는 부모에서 매 렌더 새로 생성되는 인라인 함수가 많음 → 최신값을 ref로 참조해
  // 히스토리 effect가 [open]에만 반응하도록(불필요 재실행 방지)
  const onCloseRef = useRef(onClose);
  onCloseRef.current = onClose;

  // ESC 키로 닫기
  useEffect(() => {
    if (!open) return;
    const handler = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [open, onClose]);

  // 모바일 뒤로가기 = 패널만 닫기 (페이지 이탈 방지)
  // 열릴 때 히스토리 1칸 push → popstate(뒤로가기) 시 onClose만 실행 → 리스트 화면 유지
  useEffect(() => {
    if (!open) return;
    let poppedByBack = false;
    window.history.pushState({ __slidePanel: true }, '');
    const onPop = () => { poppedByBack = true; onCloseRef.current(); };
    window.addEventListener('popstate', onPop);
    return () => {
      window.removeEventListener('popstate', onPop);
      // X·배경·ESC로 닫았고 우리가 쌓은 더미 칸이 아직 top이면 그 칸만 정리
      // (패널 안에서 router.push로 이동한 경우 state.__slidePanel 아님 → back() 안 함)
      if (!poppedByBack && (window.history.state as { __slidePanel?: boolean } | null)?.__slidePanel) {
        window.history.back();
      }
    };
  }, [open]);

  // 스크롤 잠금
  useEffect(() => {
    if (open) {
      document.body.style.overflow = 'hidden';
      return () => { document.body.style.overflow = ''; };
    }
  }, [open]);

  return (
    <>
      {/* 백드롭 */}
      <div
        className={cn(
          'fixed inset-0 z-50 bg-indigo-black/40 transition-opacity duration-200',
          open ? 'opacity-100' : 'opacity-0 pointer-events-none'
        )}
        onClick={onClose}
      />

      {/* 패널 */}
      <div
        className={cn(
          'fixed right-0 top-0 h-full z-50 w-full sm:w-96 bg-card-white shadow-2xl',
          'transition-transform duration-250 ease-out',
          open ? 'translate-x-0' : 'translate-x-full',
          className
        )}
      >
        {/* 헤더 */}
        {title && (
          <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
            <h2 className="text-sm font-bold text-indigo-black">{title}</h2>
            <button
              onClick={onClose}
              className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
            >
              <X size={18} className="text-neutral-500" />
            </button>
          </div>
        )}

        {/* 본문 */}
        <div className={cn('overflow-y-auto px-4 py-4', title ? 'h-[calc(100%-3rem)]' : 'h-full')}>
          {children}
        </div>
      </div>
    </>
  );
}
