'use client';

/**
 * CustomerNotes — 고객 메모 타임라인 (재사용)
 *   - 최종 메모 요약(맨 위 강조) + 한 줄 빠른 입력(+ 카테고리 칩) + 날짜별 타임라인
 *   - 고객상세 / 판매상세 / 대시보드 모달 등 어디서든 동일 컴포넌트 사용
 *   - customerId 없으면(고객 미연결) 안내만 표시
 */

import { useState } from 'react';
import { useCustomerNotes, useAddCustomerNote, useDeleteCustomerNote } from '@/hooks/use-customer-notes';
import { StickyNote, X, Plus } from 'lucide-react';

const CATEGORIES = ['특징', '불편', '요구', '기타'] as const;
const CAT_COLOR: Record<string, string> = {
  특징: 'bg-blue-50 text-blue-700 border-blue-200',
  불편: 'bg-rose-50 text-rose-700 border-rose-200',
  요구: 'bg-amber-50 text-amber-700 border-amber-200',
  기타: 'bg-stone-100 text-stone-600 border-stone-200',
};

function fmt(iso: string): string {
  const d = new Date(iso);
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}.${p(d.getMonth() + 1)}.${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export function CustomerNotes({ customerId, compact }: { customerId?: string | null; compact?: boolean }) {
  const { data: notes = [], isLoading } = useCustomerNotes(customerId);
  const add = useAddCustomerNote();
  const del = useDeleteCustomerNote();
  const [body, setBody] = useState('');
  const [cat, setCat] = useState<string | null>(null);

  if (!customerId) {
    return <p className="text-xs text-stone-400">고객을 연결하면 메모를 남길 수 있습니다.</p>;
  }

  const submit = () => {
    const b = body.trim();
    if (!b || add.isPending) return;
    add.mutate({ customerId, body: b, category: cat }, { onSuccess: () => { setBody(''); setCat(null); } });
  };

  const latest = notes[0];

  return (
    <div className="space-y-3">
      {/* 최종 메모 요약 */}
      {latest && (
        <div className="rounded-lg bg-stone-900 text-white px-3 py-2">
          <div className="flex items-center gap-1.5 text-[10px] text-stone-300 mb-0.5">
            <StickyNote size={11} /> 최종 메모 · {fmt(latest.created_at)}{latest.category ? ` · ${latest.category}` : ''}
          </div>
          <p className="text-sm leading-snug whitespace-pre-wrap break-words">{latest.body}</p>
        </div>
      )}

      {/* 입력 */}
      <div className="space-y-1.5">
        <div className="flex items-center gap-1">
          {CATEGORIES.map((c) => (
            <button key={c} type="button" onClick={() => setCat(cat === c ? null : c)}
              className={`text-[11px] px-2 py-0.5 rounded-full border transition ${cat === c ? CAT_COLOR[c] : 'bg-white text-stone-400 border-stone-200 hover:bg-stone-50'}`}>{c}</button>
          ))}
        </div>
        <div className="flex items-center gap-1.5">
          <input value={body} onChange={(e) => setBody(e.target.value)}
            onKeyDown={(e) => { if (e.key === 'Enter' && !e.nativeEvent.isComposing) submit(); }}
            placeholder="고객 특징·불편·요구 등 메모 입력…"
            className="flex-1 h-9 px-3 rounded-lg border border-stone-200 text-sm text-stone-800 placeholder:text-stone-400 focus:outline-none focus:border-stone-400" />
          <button onClick={submit} disabled={add.isPending || !body.trim()}
            className="h-9 px-3 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 disabled:opacity-40 flex items-center gap-1 shrink-0"><Plus size={14} />기록</button>
        </div>
      </div>

      {/* 타임라인 */}
      {isLoading ? (
        <p className="text-xs text-stone-400">불러오는 중…</p>
      ) : notes.length === 0 ? (
        <p className="text-xs text-stone-400">아직 메모가 없습니다.</p>
      ) : (
        <div className={`space-y-1.5 ${compact ? 'max-h-52 overflow-y-auto pr-1' : ''}`}>
          {notes.map((n) => (
            <div key={n.id} className="group flex items-start gap-2 rounded-lg border border-stone-100 px-2.5 py-2">
              {n.category && <span className={`text-[10px] px-1.5 py-0.5 rounded-full border shrink-0 ${CAT_COLOR[n.category] || CAT_COLOR['기타']}`}>{n.category}</span>}
              <div className="flex-1 min-w-0">
                <p className="text-sm text-stone-700 whitespace-pre-wrap break-words">{n.body}</p>
                <p className="text-[10px] text-stone-400 mt-0.5">{fmt(n.created_at)}{n.created_by ? ` · ${n.created_by}` : ''}</p>
              </div>
              <button onClick={() => del.mutate({ id: n.id, customerId })}
                className="text-stone-300 hover:text-rose-500 opacity-0 group-hover:opacity-100 transition shrink-0" aria-label="삭제"><X size={13} /></button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
