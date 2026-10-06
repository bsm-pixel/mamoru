'use client';

/**
 * 복원수리 「방문 예약 등록」 (2026-10-06)
 * 전화로 "몇 시쯤 갈게요" 한 고객을 바로 등록 → 고객 알림톡(일정 변경 링크) + 구글 캘린더 + 리마인드.
 * 충돌 = 경고 후 허용 (사장님 결정). 판정·생성은 서버(api/repair/visit-booking)가 한다.
 */

import { useEffect, useMemo, useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { X, Loader2, CalendarPlus, AlertTriangle } from 'lucide-react';
import toast from 'react-hot-toast';
import { Button } from '@/components/ui/button';
import { EscClose } from '@/components/ui/esc-close';
import { backdropClose } from '@/lib/ui/backdrop';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';

interface Slot { time: string; available: boolean; past?: boolean; conflicts?: string[] }

function kstToday(): string {
  return new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10);
}
function formatPhone(v: string): string {
  const d = v.replace(/\D/g, '').slice(0, 11);
  if (d.length < 4) return d;
  if (d.length < 8) return `${d.slice(0, 3)}-${d.slice(3)}`;
  return `${d.slice(0, 3)}-${d.slice(3, d.length - 4)}-${d.slice(-4)}`;
}

export function VisitBookingModal({ open, onClose, onCreated }: { open: boolean; onClose: () => void; onCreated?: (id: string) => void }) {
  const qc = useQueryClient();
  const [customer, setCustomer] = useState<SelectedCustomer | null>(null);
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [date, setDate] = useState(kstToday());
  const [time, setTime] = useState('');
  const [qtyM, setQtyM] = useState(0);
  const [qtyO, setQtyO] = useState(1);
  const [memo, setMemo] = useState('');
  const [slots, setSlots] = useState<Slot[]>([]);
  const [closedReason, setClosedReason] = useState<string | null>(null);
  const [loadingSlots, setLoadingSlots] = useState(false);
  const [saving, setSaving] = useState(false);
  const [confirm, setConfirm] = useState<string[] | null>(null);   // 서버가 돌려준 경고 → 확인창

  const qty = Math.max(1, qtyM + qtyO);

  // 날짜·자루 수가 바뀌면 그날 시간표 다시 읽기
  useEffect(() => {
    if (!open || !date) return;
    let alive = true;
    const run = async () => {
      setLoadingSlots(true);
      try {
        const res = await fetch(`/api/repair/visit-booking?date=${date}&qty=${qty}`);
        const j = await res.json();
        if (!alive) return;
        setSlots(j.slots || []);
        setClosedReason(j.closedReason || null);
      } catch {
        if (alive) setSlots([]);
      } finally {
        if (alive) setLoadingSlots(false);
      }
    };
    run();
    return () => { alive = false; };
  }, [open, date, qty]);

  const selectedSlot = useMemo(() => slots.find((s) => s.time === time), [slots, time]);

  function reset() {
    setCustomer(null); setName(''); setPhone(''); setDate(kstToday()); setTime('');
    setQtyM(0); setQtyO(1); setMemo(''); setConfirm(null);
  }
  function close() { reset(); onClose(); }

  async function submit(force = false) {
    if (!name.trim()) { toast.error('성함을 입력해주세요'); return; }
    if (phone.replace(/\D/g, '').length < 10) { toast.error('연락처를 확인해주세요'); return; }
    if (!time) { toast.error('방문 시간을 선택하거나 입력해주세요'); return; }
    setSaving(true);
    try {
      const res = await fetch('/api/repair/visit-booking', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name, phone, visit_date: date, visit_time: time, qty_mamoru: qtyM, qty_other: qtyO, memo, force }),
      });
      const j = await res.json().catch(() => ({}));
      if (res.status === 409 && j.needConfirm) { setConfirm(j.warnings || []); return; }
      if (!res.ok) throw new Error(j.error || '등록 실패');
      toast.success(`${j.as_id} 방문 예약을 등록했습니다${j.notified ? ' · 고객 알림톡 발송' : ''}`);
      qc.invalidateQueries({ queryKey: ['repairs'] });
      onCreated?.(j.id);
      close();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : '등록 실패');
    } finally {
      setSaving(false);
    }
  }

  if (!open) return null;

  const visible = slots.filter((s) => !s.past);
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 p-4" {...backdropClose(close)}>
      <EscClose onClose={close} />
      <div className="w-full max-w-lg max-h-[92vh] overflow-y-auto rounded-2xl bg-white shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100">
          <div className="flex items-center gap-2">
            <CalendarPlus size={18} className="text-neutral-700" />
            <h2 className="text-base font-bold text-neutral-900">방문 예약 등록</h2>
          </div>
          <button type="button" onClick={close} className="p-1 rounded text-neutral-400 hover:text-neutral-700"><X size={18} /></button>
        </div>

        <div className="px-5 py-4 space-y-5">
          {/* 1. 고객 */}
          <section>
            <p className="text-xs font-semibold text-neutral-500 mb-2">고객</p>
            <CustomerAutocomplete
              selectedCustomer={customer}
              onSelect={(c) => { setCustomer(c); setName(c.name || ''); setPhone(formatPhone(c.phone || '')); }}
              onClear={() => setCustomer(null)}
              disableInlineNewForm
            />
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2">
              <input value={name} onChange={(e) => setName(e.target.value)} placeholder="성함 *"
                className="h-10 px-3 rounded-lg border border-neutral-200 text-sm" />
              <input value={phone} onChange={(e) => setPhone(formatPhone(e.target.value))} placeholder="연락처 * (010-0000-0000)" inputMode="numeric"
                className="h-10 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
            <p className="text-[11px] text-neutral-400 mt-1">기존 고객은 위에서 검색하면 채워집니다. 처음 오시는 분은 직접 입력하면 고객으로 자동 등록돼요.</p>
          </section>

          {/* 2. 가위 수 (소요시간·일정 차단에 사용) */}
          <section>
            <p className="text-xs font-semibold text-neutral-500 mb-2">가위 수 <span className="font-normal text-neutral-400">· 모르면 그대로 두세요 (소요 {10 + (qty - 1) * 5}분)</span></p>
            <div className="flex flex-wrap gap-4">
              {([['마모루', qtyM, setQtyM], ['타사', qtyO, setQtyO]] as const).map(([label, v, set]) => (
                <div key={label} className="flex items-center gap-2">
                  <span className="text-sm text-neutral-700 w-12">{label}</span>
                  <button type="button" onClick={() => set(Math.max(0, v - 1))} className="w-8 h-8 rounded-lg border border-neutral-200 text-neutral-600">−</button>
                  <span className="w-6 text-center text-sm font-semibold">{v}</span>
                  <button type="button" onClick={() => set(v + 1)} className="w-8 h-8 rounded-lg border border-neutral-200 text-neutral-600">+</button>
                </div>
              ))}
            </div>
          </section>

          {/* 3. 날짜·시간 */}
          <section>
            <p className="text-xs font-semibold text-neutral-500 mb-2">방문 날짜 · 시간</p>
            <input type="date" value={date} onChange={(e) => { setDate(e.target.value); setTime(''); }}
              className="h-10 px-3 rounded-lg border border-neutral-200 text-sm" />
            {closedReason && (
              <p className="text-xs text-amber-700 mt-2">{closedReason === 'closed_date' ? '휴무일로 지정된 날입니다.' : '정기 휴무 요일입니다.'} 그래도 등록할 수 있어요(확인 후 등록).</p>
            )}
            <div className="mt-3">
              {loadingSlots ? (
                <p className="text-xs text-neutral-400">시간표 불러오는 중…</p>
              ) : visible.length > 0 ? (
                <div className="flex flex-wrap gap-1.5">
                  {visible.map((s) => (
                    <button key={s.time} type="button" onClick={() => setTime(s.time)}
                      title={s.conflicts?.join('\n')}
                      className={`h-8 px-2.5 rounded-lg text-xs font-semibold border transition ${
                        time === s.time ? 'bg-neutral-900 text-white border-neutral-900'
                        : s.available ? 'bg-white text-neutral-700 border-neutral-200 hover:border-neutral-400'
                        : 'bg-neutral-50 text-neutral-300 border-neutral-100 line-through'
                      }`}>
                      {s.time}
                    </button>
                  ))}
                </div>
              ) : !closedReason ? (
                <p className="text-xs text-neutral-400">이 날은 남은 시간이 없습니다. 아래에 직접 입력하세요.</p>
              ) : null}
              <div className="flex items-center gap-2 mt-2">
                <span className="text-xs text-neutral-500">직접 입력</span>
                <input type="time" value={time} onChange={(e) => setTime(e.target.value)} step={300}
                  className="h-9 px-2 rounded-lg border border-neutral-200 text-sm" />
                <span className="text-[11px] text-neutral-400">전화로 정한 시간이 칸에 없을 때</span>
              </div>
              {selectedSlot && !selectedSlot.available && selectedSlot.conflicts?.length ? (
                <p className="text-xs text-amber-700 mt-2 flex items-start gap-1"><AlertTriangle size={13} className="mt-0.5 shrink-0" />겹침: {selectedSlot.conflicts.join(' · ')}</p>
              ) : null}
            </div>
          </section>

          {/* 4. 메모 */}
          <section>
            <p className="text-xs font-semibold text-neutral-500 mb-2">메모 <span className="font-normal text-neutral-400">(선택)</span></p>
            <textarea value={memo} onChange={(e) => setMemo(e.target.value)} rows={2} placeholder="예: 오후 3시쯤 온다고 함 · 틴닝 1자루 날 나감"
              className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm resize-none" />
          </section>

          <p className="text-[11px] text-neutral-400 leading-relaxed">
            등록하면 고객에게 <b>「매장방문 접수」 알림톡</b>(일정 확인·변경 버튼 포함)이 바로 나가고, <b>구글 캘린더</b>에 기록되며 방문 전 리마인드도 자동으로 나갑니다.
          </p>
        </div>

        <div className="px-5 py-4 border-t border-neutral-100 flex justify-end gap-2">
          <Button variant="secondary" onClick={close} disabled={saving}>취소</Button>
          <Button onClick={() => submit(false)} disabled={saving}>
            {saving ? <Loader2 size={14} className="animate-spin" /> : <CalendarPlus size={14} />}
            {saving ? '등록 중…' : '예약 등록'}
          </Button>
        </div>
      </div>

      {/* 경고 후 허용 — 서버가 돌려준 경고를 보여주고 확인 시 force 로 다시 등록 */}
      {confirm && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/30 p-4" onClick={(e) => e.stopPropagation()}>
          <div className="w-full max-w-sm rounded-xl bg-white p-5 shadow-xl">
            <p className="text-sm font-bold text-neutral-900 flex items-center gap-1.5"><AlertTriangle size={16} className="text-amber-600" />확인이 필요합니다</p>
            <ul className="mt-3 space-y-1.5">
              {confirm.map((w, i) => <li key={i} className="text-sm text-neutral-700">· {w}</li>)}
            </ul>
            <p className="text-xs text-neutral-400 mt-3">그래도 이 시간으로 등록할까요? (메모에 확인 내용이 남습니다)</p>
            <div className="flex justify-end gap-2 mt-4">
              <Button variant="secondary" onClick={() => setConfirm(null)} disabled={saving}>시간 다시 고르기</Button>
              <Button onClick={() => { setConfirm(null); submit(true); }} disabled={saving}>그래도 등록</Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
