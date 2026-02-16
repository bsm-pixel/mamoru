'use client';

import { useEffect, useRef } from 'react';
import { X } from 'lucide-react';
import { cn } from '@/lib/utils/cn';

interface ModalProps {
  open: boolean;
  onClose: () => void;
  title: string;
  children: React.ReactNode;
  className?: string;
}

export function Modal({ open, onClose, title, children, className }: ModalProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);

  useEffect(() => {
    const el = dialogRef.current;
    if (!el) return;
    if (open) {
      el.showModal();
    } else {
      el.close();
    }
  }, [open]);

  if (!open) return null;

  return (
    <dialog
      ref={dialogRef}
      onClose={onClose}
      className={cn(
        'backdrop:bg-indigo-black/50 bg-card-white rounded-2xl border border-neutral-200 shadow-xl',
        'p-0 w-full max-w-lg mx-auto',
        className
      )}
    >
      {/* 헤더 */}
      <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-200">
        <h2 className="text-base font-bold text-indigo-black">{title}</h2>
        <button
          onClick={onClose}
          className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
        >
          <X size={18} className="text-neutral-500" />
        </button>
      </div>
      {/* 본문 */}
      <div className="p-5">{children}</div>
    </dialog>
  );
}
