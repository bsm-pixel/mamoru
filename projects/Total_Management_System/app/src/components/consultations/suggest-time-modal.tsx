'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { useSuggestTimes } from '@/hooks/use-consultations';
import { Plus, Trash2 } from 'lucide-react';

interface Props {
  open: boolean;
  onClose: () => void;
  consultationId: string;
}

interface TimeSlot {
  date: string;
  time: string;
}

export function SuggestTimeModal({ open, onClose, consultationId }: Props) {
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
              onChange={(e) => updateSlot(idx, 'time', e.target.value)}
              className="w-28 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
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
