'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { useHoldConsultation } from '@/hooks/use-consultations';

interface Props {
  open: boolean;
  onClose: () => void;
  consultationId: string;
}

export function HoldReasonModal({ open, onClose, consultationId }: Props) {
  const [reason, setReason] = useState('');
  const hold = useHoldConsultation();

  const handleSubmit = () => {
    if (!reason.trim()) return;
    hold.mutate(
      { id: consultationId, holdReason: reason.trim() },
      { onSuccess: () => { setReason(''); onClose(); } }
    );
  };

  return (
    <Modal open={open} onClose={onClose} title="보류 처리">
      <div className="space-y-4">
        <div>
          <label className="block text-sm font-medium text-neutral-700 mb-1">보류 사유</label>
          <textarea
            value={reason}
            onChange={(e) => setReason(e.target.value)}
            placeholder="보류 사유를 입력하세요"
            rows={3}
            className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none"
          />
        </div>
        <div className="flex justify-end gap-2">
          <Button variant="ghost" size="sm" onClick={onClose}>취소</Button>
          <Button
            size="sm"
            disabled={!reason.trim() || hold.isPending}
            onClick={handleSubmit}
          >
            {hold.isPending ? '처리 중...' : '보류 처리'}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
