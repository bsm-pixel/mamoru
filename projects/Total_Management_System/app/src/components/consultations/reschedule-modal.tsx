'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { useRescheduleConsultation } from '@/hooks/use-consultations';

interface Props {
  open: boolean;
  onClose: () => void;
  consultationId: string;
  currentDate?: string;
  currentTime?: string;
}

export function RescheduleModal({ open, onClose, consultationId, currentDate, currentTime }: Props) {
  const [date, setDate] = useState(currentDate || '');
  const [time, setTime] = useState(currentTime || '');
  const [notify, setNotify] = useState(true);
  const reschedule = useRescheduleConsultation();

  const handleSubmit = () => {
    if (!date || !time) return;
    reschedule.mutate(
      { id: consultationId, visitDate: date, visitTime: time, notify },
      { onSuccess: () => { onClose(); } }
    );
  };

  return (
    <Modal open={open} onClose={onClose} title="일정 변경">
      <div className="space-y-4">
        <div>
          <label className="block text-sm font-medium text-neutral-700 mb-1">날짜</label>
          <input
            type="date"
            value={date}
            onChange={(e) => setDate(e.target.value)}
            className="w-full h-10 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <div>
          <label className="block text-sm font-medium text-neutral-700 mb-1">시간</label>
          <input
            type="time"
            value={time}
            onChange={(e) => setTime(e.target.value)}
            className="w-full h-10 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
        </div>
        <label className="flex items-center gap-2 text-sm">
          <input
            type="checkbox"
            checked={notify}
            onChange={(e) => setNotify(e.target.checked)}
            className="rounded border-neutral-300"
          />
          변경 알림톡 발송
        </label>
        <div className="flex justify-end gap-2">
          <Button variant="ghost" size="sm" onClick={onClose}>취소</Button>
          <Button
            size="sm"
            disabled={!date || !time || reschedule.isPending}
            onClick={handleSubmit}
          >
            {reschedule.isPending ? '변경 중...' : '일정 변경'}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
