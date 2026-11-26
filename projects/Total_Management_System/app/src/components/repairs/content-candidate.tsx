'use client';

import { useState } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { useUpdateRepairFields } from '@/hooks/use-repairs';
import type { Repair } from '@/lib/supabase/types';
import { Camera, Clapperboard } from 'lucide-react';
import toast from 'react-hot-toast';
import { useQuery } from '@tanstack/react-query';
import { Modal } from '@/components/ui/modal';
import { createClient } from '@/lib/supabase/client';
import { formatDate } from '@/lib/utils/format';

/**
 * 콘텐츠 후보 (2026-11-26, 작업실 19번 · 마이그 173)
 *  마모루 콘텐츠(블로그 · 캐러셀 · 쇼츠)는 실제 복원수리 건에서 나온다. 입고 점검 때 콘텐츠로 쓸 만한 건을 표시해 두고,
 *  작업 전 · 후 촬영을 놓치지 않게 확인한다. 확인만 하고 진행은 막지 않는다. 고객 이름 · 연락처는 여기 적지 않는다.
 *  저장 = PATCH /api/repair/[id] 칸만 수정(단계가 안 바뀌므로 알림톡 없음).
 */

type ContentFields = Pick<Repair, 'id' | 'status' | 'content_candidate' | 'content_symptom' | 'content_before_shot' | 'content_after_shot' | 'content_consent'>;

/** 작업 전 촬영을 챙길 단계 — 작업에 손대기 전(입고 전후 · 작업 중) */
const BEFORE_STAGES = ['intake', 'pickup_scheduled', 'picked_up', 'inspecting', 'cost_notified', 'payment_confirmed', 'repairing'];
/** 작업 후 촬영을 챙길 단계 — 작업 중 · 작업완료(출고 전) */
const AFTER_STAGES = ['repairing', 'ready_to_ship'];

/** 지금 띄울 촬영 확인 — 화면 · 시험 공용 */
export function pendingShots(r: ContentFields): ('before' | 'after')[] {
  if (!r.content_candidate) return [];
  const out: ('before' | 'after')[] = [];
  if (!r.content_before_shot && BEFORE_STAGES.includes(r.status)) out.push('before');
  if (!r.content_after_shot && AFTER_STAGES.includes(r.status)) out.push('after');
  return out;
}

const CONSENT: { v: 'yes' | 'no' | null; label: string }[] = [
  { v: 'yes', label: '동의 ○' },
  { v: 'no', label: '거절 ✕' },
  { v: null, label: '미확인' },
];

/** 상세 화면 카드 — 후보 표시 · 증상 한 줄 · 촬영 여부 · 고객 노출 동의 */
export function ContentCandidateCard({ repair: r }: { repair: ContentFields }) {
  const update = useUpdateRepairFields();
  // 고치는 중일 때만 따로 들고 있고, 아니면 저장된 값을 그대로 보여 준다(다른 건으로 바뀌어도 옛 글이 남지 않게)
  const [draft, setDraft] = useState<{ id: string; v: string } | null>(null);
  const symptom = draft && draft.id === r.id ? draft.v : (r.content_symptom || '');
  const setSymptom = (v: string) => setDraft({ id: r.id, v });

  const save = (fields: Partial<ContentFields>) =>
    update.mutate({ id: r.id, ...fields }, { onError: (e) => toast.error(`저장 실패 — ${e.message}`) });
  const saveSymptom = () => {
    const v = symptom.trim();
    if (v !== (r.content_symptom || '')) save({ content_symptom: v || null });
    setDraft(null);
  };

  const box = (on: boolean) => `w-4 h-4 rounded-sm border flex items-center justify-center text-[10px] ${on ? 'bg-neutral-900 border-neutral-900 text-white' : 'border-neutral-300 bg-white'}`;

  return (
    <Card>
      <CardHeader className="mb-3">
        <CardTitle>
          <Clapperboard size={16} className="inline mr-1.5" />
          콘텐츠 후보
        </CardTitle>
        <button type="button" onClick={() => save({ content_candidate: !r.content_candidate })}
          className="flex items-center gap-2 text-sm font-semibold text-neutral-700" aria-pressed={r.content_candidate}>
          <span className={box(r.content_candidate)}>{r.content_candidate ? '✓' : ''}</span>
          후보로 표시
        </button>
      </CardHeader>

      {!r.content_candidate ? (
        <p className="text-xs text-neutral-400">블로그 · 인스타 · 쇼츠로 쓸 만한 수리면 입고 점검 때 표시해 두세요. 작업 전 · 후 촬영을 챙겨 드립니다.</p>
      ) : (
        <div className="space-y-3">
          <div className="flex items-center gap-3">
            <label htmlFor={`cc-sym-${r.id}`} className="text-sm text-neutral-500 w-16 shrink-0">증상</label>
            <input id={`cc-sym-${r.id}`} type="text" value={symptom} maxLength={60}
              onChange={(e) => setSymptom(e.target.value)} onBlur={saveSymptom}
              onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); (e.target as HTMLInputElement).blur(); } }}
              placeholder="한 줄 — 예: 날 끝 이빨 빠짐 / 피벗 유격"
              className="flex-1 min-w-0 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
          </div>
          <div className="flex items-center gap-3">
            <span className="text-sm text-neutral-500 w-16 shrink-0">촬영</span>
            <div className="flex gap-4 flex-wrap">
              {([['content_before_shot', '작업 전'], ['content_after_shot', '작업 후']] as const).map(([k, label]) => (
                <button key={k} type="button" onClick={() => save({ [k]: !r[k] })} className="flex items-center gap-1.5 text-sm text-neutral-700" aria-pressed={r[k]}>
                  <span className={box(r[k])}>{r[k] ? '✓' : ''}</span>{label}
                </button>
              ))}
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-sm text-neutral-500 w-16 shrink-0">노출 동의</span>
            <div className="flex rounded-lg border border-neutral-200 overflow-hidden text-xs font-semibold">
              {CONSENT.map((c) => (
                <button key={c.label} type="button" onClick={() => save({ content_consent: c.v })}
                  className={`px-3 h-8 ${r.content_consent === c.v ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500 hover:bg-neutral-50'}`}>
                  {c.label}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </Card>
  );
}

type CandidateRow = Pick<Repair, 'id' | 'as_id' | 'status' | 'received_at' | 'content_symptom' | 'content_before_shot' | 'content_after_shot' | 'content_consent'>;

/** 「콘텐츠 후보」 목록 — 접수일 · 증상 한 줄 · 촬영 여부 · 고객 노출 동의. 줄을 누르면 그 건 상세로 */
export function ContentCandidateListModal({ open, onClose, onSelect }: { open: boolean; onClose: () => void; onSelect: (id: string) => void }) {
  const { data: rows = [], isLoading } = useQuery({
    queryKey: ['repair-content-candidates'],
    enabled: open,
    staleTime: 0,
    queryFn: async () => {
      const { data, error } = await createClient().from('repairs')
        .select('id, as_id, status, received_at, content_symptom, content_before_shot, content_after_shot, content_consent')
        .eq('content_candidate', true).neq('status', 'cancelled')
        .order('received_at', { ascending: false }).limit(200);
      if (error) throw error;
      return (data || []) as CandidateRow[];
    },
  });
  const mark = (on: boolean) => <span className={on ? 'text-neutral-900 font-semibold' : 'text-neutral-300'}>{on ? '○' : '—'}</span>;
  const consent = (v: CandidateRow['content_consent']) => v === 'yes' ? <span className="font-semibold text-neutral-900">동의</span> : v === 'no' ? <span className="text-rose-600">거절</span> : <span className="text-neutral-400">미확인</span>;

  return (
    <Modal open={open} onClose={onClose} title={`콘텐츠 후보 ${rows.length ? `${rows.length}건` : ''}`} className="max-w-2xl">
      {isLoading ? (
        <p className="py-8 text-center text-sm text-neutral-400">불러오는 중…</p>
      ) : rows.length === 0 ? (
        <p className="py-8 text-center text-sm text-neutral-400">후보로 표시한 수리 건이 없습니다. 상세 화면의 「콘텐츠 후보」 카드에서 표시하세요.</p>
      ) : (
        <div className="overflow-x-auto -mx-1">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-[11px] text-neutral-400 border-b border-neutral-100">
                <th className="text-left font-semibold py-2 px-1">접수일</th>
                <th className="text-left font-semibold py-2 px-1">접수번호</th>
                <th className="text-left font-semibold py-2 px-1">증상</th>
                <th className="text-center font-semibold py-2 px-1">작업 전</th>
                <th className="text-center font-semibold py-2 px-1">작업 후</th>
                <th className="text-center font-semibold py-2 px-1">노출 동의</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.id} onClick={() => { onSelect(r.id); onClose(); }} className="border-b border-neutral-50 hover:bg-neutral-50 cursor-pointer">
                  <td className="py-2 px-1 whitespace-nowrap text-neutral-600">{formatDate(r.received_at)}</td>
                  <td className="py-2 px-1 whitespace-nowrap text-neutral-500">{r.as_id}</td>
                  <td className="py-2 px-1 min-w-[140px]">{r.content_symptom || <span className="text-neutral-300">증상 미입력</span>}</td>
                  <td className="py-2 px-1 text-center">{mark(r.content_before_shot)}</td>
                  <td className="py-2 px-1 text-center">{mark(r.content_after_shot)}</td>
                  <td className="py-2 px-1 text-center whitespace-nowrap">{consent(r.content_consent)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Modal>
  );
}

/** 작업 카드 위 촬영 확인 줄 — 후보이고 아직 안 찍었을 때만. 누르면 찍음으로 표시(진행은 막지 않음) */
export function ContentShotReminder({ repair: r }: { repair: ContentFields }) {
  const update = useUpdateRepairFields();
  const shots = pendingShots(r);
  if (!shots.length) return null;
  return (
    <div className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 space-y-1.5">
      {shots.map((s) => (
        <div key={s} className="flex items-center justify-between gap-2">
          <span className="flex items-center gap-1.5 text-[13px] font-semibold text-amber-900 min-w-0">
            <Camera size={14} className="shrink-0" />
            {s === 'before' ? '콘텐츠 후보 — 작업 전 사진 찍었나요?' : '콘텐츠 후보 — 작업 후 사진 찍었나요?'}
          </span>
          <button type="button"
            onClick={() => update.mutate({ id: r.id, [s === 'before' ? 'content_before_shot' : 'content_after_shot']: true }, { onError: (e) => toast.error(`저장 실패 — ${e.message}`) })}
            className="shrink-0 h-7 px-2.5 rounded-md bg-white border border-amber-300 text-xs font-semibold text-amber-900 hover:bg-amber-100">
            찍었어요
          </button>
        </div>
      ))}
    </div>
  );
}
