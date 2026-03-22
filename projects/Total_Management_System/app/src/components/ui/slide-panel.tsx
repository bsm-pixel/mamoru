'use client';

import { useEffect } from 'react';
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
  // ESC 키로 닫기
  useEffect(() => {
    if (!open) return;
    const handler = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [open, onClose]);

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
