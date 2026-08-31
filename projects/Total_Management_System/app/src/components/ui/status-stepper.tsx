'use client';

import { formatDate } from '@/lib/utils/format';

export interface StatusStep {
  key: string;
  label: string;
  at?: string | null;   // 해당 단계 완료 시각(있으면 라벨 아래 표시)
}

/**
 * 공용 가로 진행 스텝바 — 라이프사이클 흐름 시각화(반품·교환/주문/복원수리 등 재사용).
 *  · 완료 단계 = ✓(emerald) / 현재 = 검정 / 이후 = 회색
 *  · cancelled=true 면 취소 배너로 대체
 */
export function StatusStepper({
  steps,
  currentKey,
  cancelled,
  cancelledAt,
}: {
  steps: StatusStep[];
  currentKey: string;
  cancelled?: boolean;
  cancelledAt?: string | null;
}) {
  if (cancelled) {
    return (
      <div className="rounded-lg bg-red-50 border border-red-100 px-3 py-2 text-xs text-red-600 font-medium">
        취소됨{cancelledAt ? ` · ${formatDate(cancelledAt)}` : ''}
      </div>
    );
  }

  const found = steps.findIndex((s) => s.key === currentKey);
  const curIdx = found < 0 ? 0 : found;
  const n = steps.length;
  const cols = { gridTemplateColumns: `repeat(${n}, minmax(0,1fr))` };

  return (
    <div>
      {/* 원 + 연결선 */}
      <div className="grid" style={cols}>
        {steps.map((s, i) => {
          const done = i < curIdx;
          const cur = i === curIdx;
          return (
            <div key={s.key} className="relative flex items-center justify-center h-6">
              {i > 0 && (
                <div className={`absolute left-0 right-1/2 top-1/2 -translate-y-1/2 h-0.5 ${i <= curIdx ? 'bg-emerald-500' : 'bg-neutral-200'}`} />
              )}
              {i < n - 1 && (
                <div className={`absolute left-1/2 right-0 top-1/2 -translate-y-1/2 h-0.5 ${i < curIdx ? 'bg-emerald-500' : 'bg-neutral-200'}`} />
              )}
              <div className={`relative z-10 w-6 h-6 rounded-full flex items-center justify-center text-[11px] font-bold ${cur ? 'bg-stone-900 text-white' : done ? 'bg-emerald-500 text-white' : 'bg-neutral-200 text-neutral-400'}`}>
                {done ? '✓' : i + 1}
              </div>
            </div>
          );
        })}
      </div>
      {/* 라벨 + 날짜 */}
      <div className="grid mt-1" style={cols}>
        {steps.map((s, i) => {
          const done = i < curIdx;
          const cur = i === curIdx;
          return (
            <div key={s.key} className="text-center px-0.5">
              <div className={`text-[9.5px] leading-tight ${cur ? 'text-stone-900 font-semibold' : done ? 'text-emerald-600' : 'text-neutral-400'}`}>
                {s.label}
              </div>
              {s.at && (done || cur) && (
                <div className="text-[8.5px] text-neutral-400 mt-0.5 tabular-nums">{formatDate(s.at)}</div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
