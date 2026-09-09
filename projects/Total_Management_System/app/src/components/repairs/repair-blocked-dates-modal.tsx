'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { useSetting, useUpdateSettings } from '@/hooks/use-settings';
import { Trash2, CalendarOff, Plus } from 'lucide-react';

export interface BlockedRange {
  start: string;   // YYYY-MM-DD
  end: string;     // YYYY-MM-DD (시작과 같으면 하루)
  reason: string;
}

const KEY = 'repair.pickup_blocked_dates';

function mmdd(d: string) {
  const p = d.split('-');
  return p.length >= 3 ? `${+p[1]}/${+p[2]}` : d;
}

interface Props {
  onClose: () => void;
}

/** 수거 불가일(우체국 픽업 중지 등) 관리 — 고객 접수·수거일 변경 달력에서 선택 차단 */
export function RepairBlockedDatesModal({ onClose }: Props) {
  const saved = useSetting<BlockedRange[]>(KEY, []);
  const update = useUpdateSettings();
  const [ranges, setRanges] = useState<BlockedRange[]>(Array.isArray(saved) ? saved : []);
  const [start, setStart] = useState('');
  const [end, setEnd] = useState('');
  const [reason, setReason] = useState('');

  const add = () => {
    if (!start) return;
    const s = start;
    const e = end && end >= start ? end : start;
    setRanges((prev) => [...prev, { start: s, end: e, reason: reason.trim() }].sort((a, b) => a.start.localeCompare(b.start)));
    setStart(''); setEnd(''); setReason('');
  };
  const remove = (i: number) => setRanges((prev) => prev.filter((_, idx) => idx !== i));

  const dirty = JSON.stringify(ranges) !== JSON.stringify(saved);
  const save = () => {
    update.mutate([{ key: KEY, value: ranges }], { onSuccess: () => onClose() });
  };

  return (
    <Modal open onClose={onClose} title="수거 불가일 지정" className="max-w-lg">
      <div className="space-y-4">
        <p className="text-xs text-neutral-500 leading-relaxed">
          여기 지정한 기간은 <b>고객 접수 폼 · 수거일 변경</b> 달력에서 <b>선택 불가</b>로 표시됩니다.
          (예: 명절 우체국 픽업 중지)
        </p>

        {/* 현재 목록 */}
        <div className="space-y-1.5">
          {ranges.length === 0 ? (
            <p className="text-xs text-neutral-400 text-center py-4 border border-dashed border-neutral-200 rounded-lg">
              지정된 수거 불가일이 없습니다.
            </p>
          ) : (
            ranges.map((r, i) => (
              <div key={i} className="flex items-center gap-2 px-3 py-2 rounded-lg bg-stone-50 border border-stone-100">
                <CalendarOff size={14} className="text-red-500 shrink-0" />
                <span className="text-sm font-medium text-stone-800 tabular-nums">
                  {mmdd(r.start)}{r.end && r.end !== r.start ? ` ~ ${mmdd(r.end)}` : ''}
                </span>
                {r.reason && <span className="text-xs text-neutral-500 truncate">· {r.reason}</span>}
                <button onClick={() => remove(i)} className="ml-auto text-neutral-300 hover:text-red-500 transition" aria-label="삭제">
                  <Trash2 size={14} />
                </button>
              </div>
            ))
          )}
        </div>

        {/* 추가 폼 */}
        <div className="rounded-lg border border-neutral-200 p-3 space-y-2">
          <p className="text-[11px] font-semibold text-neutral-500">기간 추가</p>
          <div className="flex items-center gap-2 flex-wrap">
            <input type="date" value={start} onChange={(e) => setStart(e.target.value)}
              className="h-9 px-2.5 rounded-lg border border-neutral-200 bg-white text-sm" aria-label="시작일" />
            <span className="text-neutral-400 text-sm">~</span>
            <input type="date" value={end} min={start} onChange={(e) => setEnd(e.target.value)}
              className="h-9 px-2.5 rounded-lg border border-neutral-200 bg-white text-sm" aria-label="종료일(생략 시 하루)" />
          </div>
          <div className="flex items-center gap-2">
            <input type="text" value={reason} onChange={(e) => setReason(e.target.value)}
              placeholder="사유 (예: 명절 픽업 중지)"
              className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm placeholder:text-neutral-400" />
            <Button size="sm" variant="secondary" onClick={add} disabled={!start}>
              <Plus size={14} /> 추가
            </Button>
          </div>
          <p className="text-[11px] text-neutral-400">종료일을 비우면 하루만 지정됩니다.</p>
        </div>
      </div>

      <div className="flex items-center gap-2 pt-4 mt-4 border-t border-neutral-100">
        <Button className="flex-1" onClick={save} disabled={!dirty || update.isPending}>
          {update.isPending ? '저장 중…' : '저장'}
        </Button>
        <button onClick={onClose} className="px-4 text-sm text-neutral-500 hover:text-neutral-700">닫기</button>
      </div>
    </Modal>
  );
}
