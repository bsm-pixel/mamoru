'use client';

import { useState } from 'react';
import { Timer } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Modal } from '@/components/ui/modal';
import { useFieldDelay } from '@/hooks/use-consultations';

/**
 * 출장 도착 지연 안내 — 버튼 + 지연시간 선택 모달 (2026-09-22 복원)
 *
 * R1~R7 리모델(a29b989)에서 field-request-list 를 갈아엎으며 UI 만 빠졌던 기능.
 * 흐름: 버튼 → 지연 분 선택(도착 예정 미리보기) → POST /api/consultation/delay
 *       → 알림톡 field_delayed(Make 01 상담 · 💚출장 도착 지연 안내) + 상태 이력 기록
 * 사용처: 상담 상세 패널(확정 출장) · 모바일 '오늘 출장' 카드 — 한 컴포넌트로 공유
 */

const OPTIONS = [5, 10, 15, 20, 30, 45, 60];

function addMin(hhmm: string, min: number): string {
  const [h, m] = hhmm.split(':').map(Number);
  const t = h * 60 + m + min;
  return `${String(Math.floor(t / 60)).padStart(2, '0')}:${String(t % 60).padStart(2, '0')}`;
}

interface Props {
  consultationId: string;
  visitTime: string | null;
  className?: string;
  /** 좁은 카드(버튼 3개 한 줄)에선 아이콘 생략 — 360px 에서 줄바꿈 방지 */
  compact?: boolean;
}

export function FieldDelayAction({ consultationId, visitTime, className, compact }: Props) {
  const [open, setOpen] = useState(false);
  const [selected, setSelected] = useState<number | null>(null);
  const fieldDelay = useFieldDelay();
  const time = (visitTime || '').slice(0, 5);

  // 방문 시간이 없으면 '도착 예정'을 계산할 수 없어 버튼 자체를 숨긴다
  if (!/^\d{2}:\d{2}$/.test(time)) return null;

  const close = () => { setOpen(false); setSelected(null); };

  return (
    <>
      <Button
        variant="secondary"
        size="sm"
        className={`whitespace-nowrap ${className || ''}`}
        onClick={(e) => { e.stopPropagation(); setOpen(true); }}
      >
        {!compact && <Timer size={14} />}
        지연 안내
      </Button>

      <Modal open={open} onClose={close} title="출장 도착 지연 안내" className="max-w-sm">
        <div className="space-y-4" onClick={(e) => e.stopPropagation()}>
          <p className="text-xs text-neutral-500 leading-relaxed">
            고객에게 <b>도착 지연 알림톡</b>을 발송합니다. 예상 지연 시간을 선택하세요.
          </p>
          <div className="grid grid-cols-4 gap-2">
            {OPTIONS.map((min) => (
              <button
                key={min}
                type="button"
                onClick={() => setSelected(min)}
                className={`py-2 rounded-lg text-sm font-semibold transition ${
                  selected === min ? 'bg-stone-900 text-white' : 'bg-stone-100 text-neutral-600 hover:bg-stone-200'
                }`}
              >
                {min}분
              </button>
            ))}
          </div>
          {/* 고객이 받는 문구 그대로 미리보기 — 발송 전 시간 확인 */}
          <div className="rounded-lg bg-stone-50 px-3 py-2.5 text-xs text-neutral-600 space-y-0.5">
            <p>원래 시간 : <b className="text-neutral-800">{time}</b></p>
            <p>
              도착 예정 : {selected
                ? <b className="text-neutral-900">{addMin(time, selected)} (약 {selected}분 지연)</b>
                : <span className="text-neutral-400">지연 시간을 선택하세요</span>}
            </p>
          </div>
          <div className="flex gap-2">
            <Button variant="ghost" size="sm" className="flex-1" onClick={close}>취소</Button>
            <Button
              size="sm"
              className="flex-1"
              disabled={!selected || fieldDelay.isPending}
              loading={fieldDelay.isPending}
              onClick={() => selected && fieldDelay.mutate(
                { consultationId, delayMin: selected },
                { onSuccess: close },
              )}
            >
              알림톡 발송
            </Button>
          </div>
        </div>
      </Modal>
    </>
  );
}
