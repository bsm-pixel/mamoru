'use client';

/**
 * 상담 "끝난 건"·"처리 필요" 공용 조각 — 매장 방문·출장 목록이 같이 쓴다 (2026-11-09)
 *
 * 지난내역 기준 = 상태가 끝난 건(완료·취소), 방문일 상관없음.
 *   전엔 매장 방문만 "방문일이 지난 건"으로 나눠, 방문일 전에 미리 취소된 건이 어느 탭에도 안 보였다.
 * 방문일이 지났는데 아직 확정인 건 = 할 일 → 확정 탭 맨 위 「처리 필요」로 올린다(지난내역에 섞지 않음).
 */
import type { ReactNode } from 'react';
import { useConsultations } from '@/hooks/use-consultations';
import type { Consultation, ConsultationHistory } from '@/lib/supabase/types';
import { formatPhone } from '@/lib/utils/format';

export type PastFilter = 'all' | 'completed' | 'cancelled';

/** 지난내역 조회 조건 — 칩 선택에 따라 */
export function pastStatuses(f: PastFilter): string[] {
  return f === 'all' ? ['completed', 'cancelled'] : [f];
}

/** 이력 한 줄의 사유 문구. 고객 취소는 note에 "고객 취소: …"로 저장, 관리자 취소는 note 없이 changed_by만 있다 */
export function historyNoteText(h: Pick<ConsultationHistory, 'note' | 'to_status' | 'changed_by'>): string | null {
  if (h.note) return h.note;
  if (h.to_status === 'cancelled' && h.changed_by) return '관리자 취소';
  return null;
}

/** "11/9" */
function md(iso: string): string {
  const d = new Date(iso);
  return `${d.getMonth() + 1}/${d.getDate()}`;
}

/** 지난내역 행 보조줄: 처리일 · 예약일 + 사유 */
export function ClosedMeta({ c, h }: { c: Consultation; h?: ConsultationHistory }) {
  const processedAt = h?.created_at || c.updated_at;
  const reason = h ? historyNoteText(h) : null;
  return (
    <>
      <div className="flex items-center gap-2 mt-1 text-xs text-neutral-500">
        <span>{formatPhone(c.phone)}</span>
        <span className="font-medium text-neutral-600">{md(processedAt)} {c.status === 'cancelled' ? '취소' : '완료'}</span>
        {c.visit_date && <span>· 예약 {c.visit_date.slice(5).replace('-', '/')} {c.visit_time || ''}</span>}
      </div>
      {reason && (
        <p className="mt-0.5 text-xs text-neutral-500 line-clamp-1" title={reason}>└ {reason}</p>
      )}
    </>
  );
}

/** 지난내역 칩: 전체 · 완료 · 취소 */
export function PastFilterChips({ value, onChange }: { value: PastFilter; onChange: (v: PastFilter) => void }) {
  const items: { key: PastFilter; label: string }[] = [
    { key: 'all', label: '전체' },
    { key: 'completed', label: '완료' },
    { key: 'cancelled', label: '취소' },
  ];
  return (
    <div className="flex items-center gap-1.5">
      {items.map((it) => (
        <button
          key={it.key}
          onClick={() => onChange(it.key)}
          className={`shrink-0 px-2.5 py-1 rounded-full text-[11px] font-semibold transition ${
            value === it.key ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
          }`}
        >
          {it.label}
        </button>
      ))}
    </div>
  );
}

/** 방문일이 지났는데 아직 「확정」인 건 — 건수만 (단계 카드 숫자용) */
export function useOverdueCount(type: 'store_visit' | 'field_request') {
  const { data } = useConsultations({ status: 'confirmed', dateFilter: 'past', type, limit: 1 });
  return data?.total ?? 0;
}

/** 확정 탭 맨 위 「처리 필요」 묶음 — 있을 때만 */
export function OverdueBlock({
  type,
  search,
  renderRow,
}: {
  type: 'store_visit' | 'field_request';
  search: string;
  renderRow: (c: Consultation) => ReactNode;
}) {
  const { data } = useConsultations({
    status: 'confirmed',
    dateFilter: 'past',
    type,
    search,
    limit: 50,
    orderBy: 'visit_date_asc',
  });
  const rows = data?.consultations || [];
  if (rows.length === 0) return null;
  return (
    <div className="border-b border-neutral-200">
      <div className="px-4 py-2 bg-amber-50 text-xs font-bold text-amber-700">
        처리 필요 <span className="ml-1 font-normal">방문일이 지났는데 완료·취소 처리가 안 된 건 {rows.length}건</span>
      </div>
      <div className="divide-y divide-neutral-100">{rows.map(renderRow)}</div>
    </div>
  );
}
