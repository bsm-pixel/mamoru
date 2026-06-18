'use client';

import { useEffect, useState } from 'react';
import { X, Search, Loader2, AlertTriangle, ArrowRight, Check } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { formatKRW, formatPhone } from '@/lib/utils/format';
import { useMergeCustomers } from '@/hooks/use-customers';
import toast from 'react-hot-toast';

interface Hit {
  id: string;
  name: string;
  phone: string | null;
  company_name?: string | null;
  customer_type?: string | null;
}

interface VictimSummary {
  sales: number;
  repairs: number;
  consultations: number;
  contracts: number;
  outstanding: number;
}

interface Props {
  open: boolean;
  onClose: () => void;
  /** 거래를 흡수해 유지할 주 고객 (현재 보고 있는 고객) */
  primary: { id: string; name: string; phone?: string | null };
  onMerged?: () => void;
}

export function CustomerMergeModal({ open, onClose, primary, onMerged }: Props) {
  const [q, setQ] = useState('');
  const [results, setResults] = useState<Hit[]>([]);
  const [searching, setSearching] = useState(false);
  const [selected, setSelected] = useState<Record<string, Hit>>({});
  const [summaries, setSummaries] = useState<Record<string, VictimSummary>>({});
  const [confirm, setConfirm] = useState(false);
  const merge = useMergeCustomers();

  // 모달 닫힐 때 초기화
  useEffect(() => {
    if (!open) { setQ(''); setResults([]); setSelected({}); setSummaries({}); setConfirm(false); }
  }, [open]);

  // 검색 (디바운스)
  useEffect(() => {
    if (!open) return;
    if (q.trim().length < 2) { setResults([]); return; }
    let active = true;
    setSearching(true);
    const t = setTimeout(async () => {
      try {
        const res = await fetch(`/api/customers/search?q=${encodeURIComponent(q.trim())}`);
        const json = await res.json();
        if (active) setResults(((json.customers || []) as Hit[]).filter((c) => c.id !== primary.id));
      } finally {
        if (active) setSearching(false);
      }
    }, 250);
    return () => { active = false; clearTimeout(t); };
  }, [q, open, primary.id]);

  if (!open) return null;

  const victims = Object.values(selected);

  async function toggle(c: Hit) {
    setSelected((s) => {
      const next = { ...s };
      if (next[c.id]) delete next[c.id];
      else next[c.id] = c;
      return next;
    });
    // 미리보기 카운트 로드 (선택 시 1회)
    if (!selected[c.id] && !summaries[c.id]) {
      try {
        const res = await fetch(`/api/customers/${c.id}`);
        const d = await res.json();
        setSummaries((m) => ({
          ...m,
          [c.id]: {
            sales: d.sales?.length || 0,
            repairs: d.repairs?.length || 0,
            consultations: d.consultations?.length || 0,
            contracts: d.contracts?.length || 0,
            outstanding: d.customer?.outstanding_balance || 0,
          },
        }));
      } catch { /* 미리보기 실패는 무시 (병합 자체엔 영향 없음) */ }
    }
  }

  async function doMerge() {
    try {
      await merge.mutateAsync({ primaryId: primary.id, victimIds: victims.map((v) => v.id) });
      toast.success(`${victims.length}명을 ${primary.name}(으)로 병합했습니다`);
      onMerged?.();
      onClose();
    } catch (e) {
      toast.error((e as Error).message || '병합 실패');
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-[460px] max-h-[85vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        {/* 헤더 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">고객 병합</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="flex-1 overflow-y-auto p-4 space-y-3">
          {/* 주 고객 안내 */}
          <div className="rounded-lg bg-stone-50 border border-stone-200 p-3 text-xs">
            <p className="text-neutral-500 mb-0.5">유지할 주 고객 (모든 거래가 이쪽으로 통합됩니다)</p>
            <p className="text-sm font-bold text-stone-900">{primary.name} <span className="font-normal text-neutral-400">{formatPhone(primary.phone || '')}</span></p>
          </div>

          {!confirm ? (
            <>
              {/* 검색 */}
              <div className="relative">
                <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
                <input
                  type="text"
                  value={q}
                  onChange={(e) => setQ(e.target.value)}
                  placeholder="합칠 중복 고객 검색 (이름·전화)"
                  className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400"
                  autoFocus
                />
              </div>

              {/* 검색 결과 */}
              <div className="space-y-1">
                {searching && <div className="flex justify-center py-3"><Loader2 size={16} className="animate-spin text-neutral-400" /></div>}
                {!searching && q.trim().length >= 2 && results.length === 0 && (
                  <p className="text-xs text-neutral-400 text-center py-3">검색 결과 없음</p>
                )}
                {results.map((c) => {
                  const on = !!selected[c.id];
                  return (
                    <button
                      key={c.id}
                      onClick={() => toggle(c)}
                      className={`w-full flex items-center gap-2 px-3 py-2 rounded-lg border text-left transition ${on ? 'border-stone-900 bg-stone-50' : 'border-neutral-200 hover:bg-neutral-50'}`}
                    >
                      <span className={`w-4 h-4 rounded flex items-center justify-center shrink-0 ${on ? 'bg-stone-900 text-white' : 'border border-neutral-300'}`}>
                        {on && <Check size={11} />}
                      </span>
                      <span className="flex-1 min-w-0">
                        <span className="text-sm font-medium text-stone-900">{c.name}</span>
                        <span className="text-xs text-neutral-400 ml-1.5">{formatPhone(c.phone || '')}</span>
                        {c.company_name && <span className="text-xs text-neutral-400 ml-1.5 truncate">· {c.company_name}</span>}
                      </span>
                    </button>
                  );
                })}
              </div>

              {/* 선택된 흡수 대상 미리보기 */}
              {victims.length > 0 && (
                <div className="rounded-lg border border-amber-200 bg-amber-50/50 p-3 space-y-1.5">
                  <p className="text-xs font-semibold text-amber-700">흡수 대상 {victims.length}명 → {primary.name}</p>
                  {victims.map((v) => {
                    const s = summaries[v.id];
                    return (
                      <div key={v.id} className="flex items-center gap-2 text-xs">
                        <ArrowRight size={11} className="text-amber-500 shrink-0" />
                        <span className="font-medium text-stone-800">{v.name}</span>
                        <span className="text-neutral-400">{formatPhone(v.phone || '')}</span>
                        {s && (
                          <span className="ml-auto text-[11px] text-neutral-500 shrink-0">
                            판매 {s.sales} · 수리 {s.repairs}{s.outstanding > 0 ? ` · 미수 ${formatKRW(s.outstanding)}` : ''}
                          </span>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </>
          ) : (
            /* 확인 단계 */
            <div className="rounded-lg border border-red-200 bg-red-50/60 p-3 space-y-2">
              <div className="flex items-center gap-2 text-red-600">
                <AlertTriangle size={16} />
                <p className="text-sm font-bold">병합 확인</p>
              </div>
              <p className="text-xs text-neutral-600 leading-relaxed">
                아래 {victims.length}명의 <b>모든 판매·납품·복원수리·상담·계약·주문</b>이
                <b> {primary.name}</b>(으)로 이관되고, 흡수된 고객은 목록에서 숨겨집니다(이력은 보존).
                매출·송장 표시 이름도 {primary.name}(으)로 통일됩니다.
              </p>
              <ul className="text-xs text-stone-700 space-y-0.5">
                {victims.map((v) => (
                  <li key={v.id} className="flex items-center gap-1.5">
                    <ArrowRight size={11} className="text-red-400" />{v.name} {formatPhone(v.phone || '')}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>

        {/* 하단 액션 */}
        <div className="px-4 py-3 border-t border-neutral-200 flex gap-2">
          {!confirm ? (
            <Button className="flex-1" disabled={victims.length === 0} onClick={() => setConfirm(true)}>
              다음 ({victims.length}명 선택)
            </Button>
          ) : (
            <>
              <Button variant="ghost" onClick={() => setConfirm(false)} disabled={merge.isPending}>뒤로</Button>
              <Button className="flex-1 bg-red-600 hover:bg-red-700" onClick={doMerge} disabled={merge.isPending}>
                {merge.isPending ? '병합 중...' : '병합 실행'}
              </Button>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
