'use client';

/**
 * 수동 일정 확정 모달 (079)
 *
 * 사장님이 DM/유선으로 협의된 일정을 즉시 확정 처리.
 * 기존 접수 건의 상태 + visit_date + visit_time 한 번에 변경 → /api/consultation/[id] PATCH.
 *
 * 자동 처리 (기존 PATCH 로직 재사용):
 *   1. 상태 전이 검증 (transitions.ts)
 *   2. consultation_history 이력 기록
 *   3. Google Calendar 동기화 (이전 상태 이벤트 정리 + confirmed 이벤트 생성)
 *   4. 알림톡 자동 발송 (confirmed → field_confirmed/confirmed 템플릿)
 *
 * 사장님 룰: 사장님 측 흐름이라 closed_dates 검증 X (항상 유동)
 */

import { useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { Info, AlertCircle } from 'lucide-react';
import toast from 'react-hot-toast';
import { CONSULTATION_STATUS_LABEL } from '@/lib/utils/format';

interface Props {
  open: boolean;
  onClose: () => void;
  consultationId: string;
  customerName: string;
  consultationType: 'store_visit' | 'field_request' | 'talk_consult';
  currentStatus: string;
  currentVisitDate?: string | null;
  currentVisitTime?: string | null;
}

export function ManualConfirmModal({
  open,
  onClose,
  consultationId,
  customerName,
  consultationType,
  currentStatus,
  currentVisitDate,
  currentVisitTime,
}: Props) {
  const queryClient = useQueryClient();
  const [date, setDate] = useState(currentVisitDate || '');
  const [time, setTime] = useState(currentVisitTime || '');
  const [note, setNote] = useState('');
  const [submitting, setSubmitting] = useState(false);

  const typeLabel = consultationType === 'field_request' ? '출장요청' : '매장방문';
  const statusLabel = CONSULTATION_STATUS_LABEL[currentStatus] || currentStatus;
  const isAlreadyConfirmed = currentStatus === 'confirmed';

  const handleClose = () => {
    if (submitting) return;
    setDate(currentVisitDate || '');
    setTime(currentVisitTime || '');
    setNote('');
    onClose();
  };

  const handleSubmit = async () => {
    if (!date || !time) {
      toast.error('날짜와 시간을 모두 입력해주세요');
      return;
    }

    setSubmitting(true);
    try {
      const res = await fetch(`/api/consultation/${consultationId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          status: 'confirmed',
          visit_date: date,
          visit_time: time,
          note: note.trim() || '수동 일정 확정 (외부 협의)',
        }),
      });

      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '확정 처리 실패');
      }

      toast.success('일정이 확정되었습니다 — 알림톡 + 캘린더 자동 처리');
      queryClient.invalidateQueries({ queryKey: ['consultation', consultationId] });
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      handleClose();
    } catch (err) {
      toast.error('확정 실패: ' + (err instanceof Error ? err.message : String(err)));
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Modal open={open} onClose={handleClose} title="수동 일정 확정">
      <div className="space-y-4">
        {/* 안내 */}
        <div className="flex items-start gap-2 p-3 rounded-lg bg-blue-50 text-xs text-blue-700">
          <Info size={14} className="shrink-0 mt-0.5" />
          <div className="space-y-1">
            <p>DM/유선 협의된 일정을 즉시 확정합니다.</p>
            <p className="text-blue-600">
              자동 처리: <span className="font-medium">알림톡 발송 + Google Calendar 갱신 + 이력 기록</span>
            </p>
          </div>
        </div>

        {/* 현재 상태 */}
        <div className="grid grid-cols-2 gap-2 text-xs">
          <div className="p-2 rounded-lg bg-neutral-50">
            <p className="text-neutral-500 mb-1">고객</p>
            <p className="font-bold text-neutral-800">{customerName}</p>
          </div>
          <div className="p-2 rounded-lg bg-neutral-50">
            <p className="text-neutral-500 mb-1">유형 / 현재 상태</p>
            <p className="font-medium text-neutral-800">
              {typeLabel} <span className="text-neutral-400">·</span> {statusLabel}
            </p>
          </div>
        </div>

        {isAlreadyConfirmed && (
          <div className="flex items-start gap-2 p-2 rounded-lg bg-amber-50 text-xs text-amber-700">
            <AlertCircle size={14} className="shrink-0 mt-0.5" />
            <span>이미 확정된 건입니다. 일정만 변경됩니다 — 고객에게 변경 알림톡이 재발송됩니다.</span>
          </div>
        )}

        {/* 날짜/시간 입력 */}
        <div className="grid grid-cols-2 gap-2">
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">날짜</label>
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
            />
          </div>
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">시간</label>
            <input
              type="time"
              value={time}
              step="600"
              onChange={(e) => setTime(e.target.value)}
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
            />
          </div>
        </div>

        {/* 메모 (선택) */}
        <div>
          <label className="text-xs text-neutral-500 mb-1 block">메모 (선택)</label>
          <input
            type="text"
            value={note}
            onChange={(e) => setNote(e.target.value)}
            placeholder="예: DM 협의 / 유선 협의 / 사장님 직접 확정"
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400"
          />
        </div>

        {/* 액션 */}
        <div className="flex justify-end gap-2 pt-2">
          <Button variant="ghost" size="sm" onClick={handleClose} disabled={submitting}>
            취소
          </Button>
          <Button size="sm" onClick={handleSubmit} disabled={submitting || !date || !time}>
            {submitting ? '처리 중...' : '확정 처리'}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
