'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { useSuggestTimes } from '@/hooks/use-consultations';
import { Plus, Trash2, Info } from 'lucide-react';

interface Props {
  open: boolean;
  onClose: () => void;
  consultationId: string;
  prefDays?: string[];    // 고객 가능요일 (예: ["월", "수", "금"])
  prefTimes?: string[];   // 고객 선호시간대 (예: ["오전", "오후"])
}

interface TimeSlot {
  date: string;
  time: string;
}

export function SuggestTimeModal({ open, onClose, consultationId, prefDays, prefTimes }: Props) {
  const [slots, setSlots] = useState<TimeSlot[]>([{ date: '', time: '' }]);
  const suggest = useSuggestTimes();

  const addSlot = () => {
    if (slots.length >= 3) return;
    setSlots([...slots, { date: '', time: '' }]);
  };

  const removeSlot = (idx: number) => {
    setSlots(slots.filter((_, i) => i !== idx));
  };

  const updateSlot = (idx: number, field: keyof TimeSlot, value: string) => {
    const next = [...slots];
    next[idx] = { ...next[idx], [field]: value };
    setSlots(next);
  };

  const validSlots = slots.filter((s) => s.date && s.time);

  const handleSubmit = () => {
    if (validSlots.length === 0) return;
    suggest.mutate(
      { consultationId, suggestions: validSlots },
      {
        onSuccess: () => {
          setSlots([{ date: '', time: '' }]);
          onClose();
        },
      }
    );
  };

  return (
    <Modal open={open} onClose={onClose} title="시간 제안">
      <div className="space-y-4">
        {/* 고객 가능요일/선호시간대 참고 */}
        {((prefDays && prefDays.length > 0) || (prefTimes && prefTimes.length > 0)) && (
          <div className="flex items-start gap-2 p-2.5 rounded-lg bg-blue-50 text-xs text-blue-700">
            <Info size={14} className="shrink-0 mt-0.5" />
            <div className="space-y-1">
              {prefDays && prefDays.length > 0 && (
                <div className="flex items-center gap-1 flex-wrap">
                  <span className="font-semibold">가능요일:</span>
                  {prefDays.map(d => (
                    <span key={d} className="px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 font-medium">{d}</span>
                  ))}
                </div>
              )}
              {prefTimes && prefTimes.length > 0 && (
                <div className="flex items-center gap-1 flex-wrap">
                  <span className="font-semibold">선호시간:</span>
                  {prefTimes.map(t => (
                    <span key={t} className="px-1.5 py-0.5 rounded bg-amber-100 text-amber-700 font-medium">{t}</span>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {slots.map((slot, idx) => (
          <div key={idx} className="flex items-center gap-2">
            <input
              type="date"
              value={slot.date}
              onChange={(e) => updateSlot(idx, 'date', e.target.value)}
              className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
            <input
              type="time"
              value={slot.time}
              step="600"
              onChange={(e) => updateSlot(idx, 'time', e.target.value)}
              className="w-32 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
            />
            {slots.length > 1 && (
              <button
                onClick={() => removeSlot(idx)}
                className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-error-soft transition text-neutral-400 hover:text-error"
              >
                <Trash2 size={14} />
              </button>
            )}
          </div>
        ))}

        {slots.length < 3 && (
          <button
            onClick={addSlot}
            className="flex items-center gap-1 text-xs text-terracotta hover:underline"
          >
            <Plus size={14} /> 시간 추가
          </button>
        )}

        <div className="flex justify-end gap-2 pt-2">
          <Button variant="ghost" size="sm" onClick={onClose}>취소</Button>
          <Button
            size="sm"
            disabled={validSlots.length === 0 || suggest.isPending}
            onClick={handleSubmit}
          >
            {suggest.isPending ? '전송 중...' : `${validSlots.length}건 제안`}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
