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
  consultationType?: string; // 'store_visit' | 'field_request'
  uniqueId?: string;
}

// 영업시간 가드 — consultation_settings start_hour(10)/end_hour(20) 기준. 오전/오후 오입력으로 생기는 '유령 예약'(슬롯 미차단) 방지
const OPEN = '10:00';
const CLOSE = '20:00';

export function RescheduleModal({ open, onClose, consultationId, currentDate, currentTime, consultationType, uniqueId }: Props) {
  const [date, setDate] = useState(currentDate || '');
  const [time, setTime] = useState(currentTime || '');
  const [notify, setNotify] = useState(true);
  const [timeError, setTimeError] = useState(''); // 영업시간 밖 경고
  const reschedule = useRescheduleConsultation();

  const handleSubmit = () => {
    if (!date || !time) return;
    // 영업시간 밖 시간 차단 (04:30 같은 오전/오후 오입력 방지 — 이 시간은 어떤 슬롯도 못 막음)
    if (time < OPEN || time >= CLOSE) {
      setTimeError(`영업시간(${OPEN}~${CLOSE}) 안의 시간을 선택하세요. 오전/오후를 확인하세요.`);
      return;
    }
    setTimeError('');
    reschedule.mutate(
      {
        id: consultationId,
        visitDate: date,
        visitTime: time,
        consultationType,
        uniqueId,
        notify,
      },
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
            min={OPEN}
            max={CLOSE}
            onChange={(e) => { setTime(e.target.value); if (timeError) setTimeError(''); }}
            className="w-full h-10 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          />
          {timeError && <p className="mt-1 text-xs text-red-500">{timeError}</p>}
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
