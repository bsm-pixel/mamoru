'use client';

import { useState } from 'react';
import { Modal } from './modal';
import { Button } from './button';

interface ConfirmModalProps {
  open: boolean;
  onClose: () => void;
  onConfirm: () => void | Promise<void>;
  title: string;
  message: string | React.ReactNode;
  confirmLabel?: string;
  cancelLabel?: string;
  variant?: 'default' | 'danger';
  loading?: boolean;
}

/**
 * 확인 모달 — 위험 액션 실행 전 사용자 확인용
 * variant='danger': 빨간 확인 버튼 (삭제, 취소 등)
 */
export function ConfirmModal({
  open,
  onClose,
  onConfirm,
  title,
  message,
  confirmLabel = '확인',
  cancelLabel = '닫기', // 2026-11-15 「취소」는 '취소하는 창'에서 어느 쪽인지 헷갈려 「닫기」로
  variant = 'default',
  loading: externalLoading,
}: ConfirmModalProps) {
  const [internalLoading, setInternalLoading] = useState(false);
  const isLoading = externalLoading ?? internalLoading;

  const handleConfirm = async () => {
    setInternalLoading(true);
    try {
      await onConfirm();
      onClose();
    } catch {
      // 에러는 onConfirm 내부에서 toast 등으로 처리
    } finally {
      setInternalLoading(false);
    }
  };

  return (
    <Modal open={open} onClose={onClose} title={title} className="max-w-sm">
      <div className="space-y-4">
        <div className="text-sm text-neutral-600">{message}</div>
        <div className="flex gap-2 justify-end">
          {/* 2026-11-15 확인 창은 커서를 안전한 쪽(닫기·취소)에 둔다 — lib/ui/modal-keys.ts 가 data-safe 를 먼저 찾는다 */}
          <Button variant="ghost" size="sm" onClick={onClose} disabled={isLoading} data-safe>
            {cancelLabel}
          </Button>
          <Button
            variant={variant === 'danger' ? 'danger' : 'primary'}
            size="sm"
            onClick={handleConfirm}
            loading={isLoading}
          >
            {confirmLabel}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
