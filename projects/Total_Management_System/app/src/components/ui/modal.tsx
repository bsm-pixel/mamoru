'use client';

import { useEffect, useRef, useState } from 'react';
import { X } from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { attachModalKeys, type ModalKeys } from '@/lib/ui/modal-keys';
import { DiscardAsk } from './discard-ask';

interface ModalProps {
  open: boolean;
  onClose: () => void;
  title: string;
  children: React.ReactNode;
  className?: string;
  /**
   * 모바일 카메라 촬영(capture) 시 네이티브 <dialog> 가 포커스 손실로 임의 close 되는 버그 방지.
   * true 면 의도된 닫기(X 버튼/부모 open=false)만 허용하고, 예기치 않은 close 는 재오픈한다.
   * Escape 닫기는 비활성(검수처럼 작성 중 데이터 보호가 필요한 모달에만 사용).
   */
  preventAutoClose?: boolean;
}

export function Modal({ open, onClose, title, children, className, preventAutoClose }: ModalProps) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  // 의도된 닫기(X 버튼 클릭 또는 부모가 open=false) 여부 추적
  const closingRef = useRef(false);

  useEffect(() => {
    const el = dialogRef.current;
    if (!el) return;
    if (open) {
      if (!el.open) el.showModal();
    } else {
      closingRef.current = true;
      el.close();
    }
  }, [open]);

  // 2026-11-15 입력 창 공통 키보드 동작 — 첫 입력 칸으로 초점, Ctrl+Enter 저장, Enter 다음 칸, 입력 여부 기억
  //   (Tab 가두기·닫힐 때 초점 되돌리기는 <dialog> 가 이미 한다)
  const keys = useRef<ModalKeys | null>(null);
  const [asking, setAsking] = useState(false);
  useEffect(() => {
    const el = dialogRef.current;
    if (!open || !el) return;
    keys.current = attachModalKeys(el, { trapTab: false, isTop: () => !el.querySelector('dialog[open]') });
    return () => { keys.current?.dispose(); keys.current = null; setAsking(false); };
  }, [open]);

  if (!open) return null;

  const requestClose = () => {
    closingRef.current = true;
    onClose();
  };

  return (
    <dialog
      ref={dialogRef}
      onClose={() => {
        // preventAutoClose: 의도하지 않은 close(카메라/포커스 손실)면 재오픈
        if (preventAutoClose && !closingRef.current) {
          dialogRef.current?.showModal();
          return;
        }
        closingRef.current = false;
        onClose();
      }}
      onCancel={(e) => {
        // preventAutoClose: Escape 로 인한 닫힘 방지 (X 버튼으로만 닫기)
        if (preventAutoClose) { e.preventDefault(); return; }
        // 입력한 내용이 있으면 브라우저가 창을 그냥 닫지 못하게 (Esc 는 아래 onKeyDown 이 처리)
        if (asking || keys.current?.isDirty()) e.preventDefault();
      }}
      onKeyDown={(e) => {
        // 2026-11-15 Esc 를 여기서 직접 받는다 — 브라우저 기본 닫기는 Esc 를 연달아 누르면 막을 수 없어 입력이 날아간다
        if (e.key !== 'Escape' || e.nativeEvent.isComposing) return;
        e.preventDefault();
        if (preventAutoClose) return;
        if (asking) { setAsking(false); return; }           // 묻는 중의 Esc = 계속 입력
        if (keys.current?.isDirty()) { setAsking(true); return; }
        requestClose();
      }}
      className={cn(
        'backdrop:bg-indigo-black/50 bg-card-white rounded-2xl border border-neutral-200 shadow-xl',
        'p-0 w-full max-w-lg mx-auto',
        className
      )}
    >
      {/* 헤더 */}
      <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-200">
        <h2 className="text-base font-bold text-indigo-black">{title}</h2>
        <button title="닫기" aria-label="닫기"
          onClick={requestClose}
          className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-warm-ivory transition"
        >
          <X size={18} className="text-neutral-500" />
        </button>
      </div>
      {/* 본문 */}
      <div className="p-5">{children}</div>
      {asking && <DiscardAsk onKeep={() => setAsking(false)} onDiscard={() => { setAsking(false); requestClose(); }} />}
    </dialog>
  );
}
