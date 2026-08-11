'use client';

/**
 * DashboardCalendarPanel — 대시보드 전용 컴팩트 달력 + 선택일 타임라인 (2분할)
 *
 * 시안 B+ (2026-05-26 사장님 채택):
 *   - 좌측: 월간 달력 (매장/출장/수리 3색 점 표시)
 *   - 우측: 선택일 일정 타임라인 (기본값 = 오늘)
 *   - 트렌디 디자인 (rounded-2xl, stone 베이스, 절제된 상태색)
 *
 * 데이터:
 *   - useConsultations × 2 (store_visit, field_request)
 *   - useRepairSchedule (복원수리 직접방문)
 *   - 모두 staleTime 30s 캐싱 (use-consultations.ts L34)
 */

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import { useConsultations } from '@/hooks/use-consultations';
import { CustomerNotes } from '@/components/shared/customer-notes';
import { activityDisplay } from '@/lib/customer/display';
import { useRepairSchedule, type RepairScheduleItem } from '@/hooks/use-repairs';
import { formatPhone, CONSULTATION_STATUS_LABEL, CONSULTATION_STATUS_COLOR } from '@/lib/utils/format';
import { ChevronLeft, ChevronRight, X, Calendar as CalendarIcon, Phone, MapPin, ArrowRight, StickyNote } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

type CalendarEvent =
  | { kind: 'consult'; data: Consultation }
  | { kind: 'repair';  data: RepairScheduleItem };

const DAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];
const CTYPE_LABEL: Record<string, string> = { store_visit: '매장', field_request: '출장', talk_consult: '톡상담' };

interface ConsultContext {
  company: string | null;
  pastConsults: { id: string; consultation_type: string; visit_date: string | null; visit_time: string | null; status: string }[];
  sales: { id: string; sale_number: string; sale_date: string; total_amount: number }[];
}

/** 모달용 컨텍스트: 매장명(고객) + 이전 상담 + 구매 이력 (customer_id 우선, 없으면 전화번호) */
async function fetchConsultContext(c: Consultation): Promise<ConsultContext> {
  const sb = createClient();
  const cid = c.customer_id;
  const pn = c.phone_normalized;

  let company: string | null = null;
  if (cid) {
    const { data } = await sb.from('customers').select('company_name').eq('id', cid).single();
    company = (data as { company_name: string | null } | null)?.company_name ?? null;
  }

  let pastConsults: ConsultContext['pastConsults'] = [];
  if (cid || pn) {
    let q = sb.from('consultations')
      .select('id, consultation_type, visit_date, visit_time, status')
      .neq('id', c.id)
      .order('visit_date', { ascending: false, nullsFirst: false })
      .limit(5);
    q = cid ? q.eq('customer_id', cid) : q.eq('phone_normalized', pn as string);
    const { data } = await q;
    pastConsults = data ?? [];
  }

  let sales: ConsultContext['sales'] = [];
  if (cid) {
    const { data } = await sb.from('offline_sales')
      .select('id, sale_number, sale_date, total_amount')
      .eq('customer_id', cid)
      .is('cancelled_at', null)
      .order('sale_date', { ascending: false })
      .limit(5);
    sales = data ?? [];
  }

  return { company, pastConsults, sales };
}

function getCalendarDays(year: number, month: number) {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const startOffset = firstDay.getDay();
  const totalDays = lastDay.getDate();
  const days: (number | null)[] = [];
  for (let i = 0; i < startOffset; i++) days.push(null);
  for (let d = 1; d <= totalDays; d++) days.push(d);
  while (days.length % 7 !== 0) days.push(null);
  return days;
}

function formatYYYYMMDD(year: number, month: number, day: number): string {
  return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

export function DashboardCalendarPanel() {
  const router = useRouter();
  const today = new Date();
  const [year, setYear] = useState(today.getFullYear());
  const [month, setMonth] = useState(today.getMonth());
  const todayStr = formatYYYYMMDD(today.getFullYear(), today.getMonth(), today.getDate());
  // 기본값 = 오늘
  const [selectedDate, setSelectedDate] = useState<string>(todayStr);
  // 일정 클릭 시 상세 모달 (상세페이지 이동 대신)
  const [detail, setDetail] = useState<Consultation | null>(null);
  const { data: ctx, isLoading: ctxLoading } = useQuery({
    queryKey: ['consult-context', detail?.id],
    enabled: !!detail,
    staleTime: 60_000,
    queryFn: () => fetchConsultContext(detail as Consultation),
  });

  const monthStart = formatYYYYMMDD(year, month, 1);
  const monthEnd = formatYYYYMMDD(year, month, new Date(year, month + 1, 0).getDate());

  const { data: storeData } = useConsultations({
    statuses: ['confirmed', 'in_progress'],
    type: 'store_visit',
    limit: 200,
    dateFilter: 'all',
    orderBy: 'visit_date_asc',
  });
  const { data: fieldData } = useConsultations({
    statuses: ['confirmed', 'suggested', 'in_progress'],
    type: 'field_request',
    limit: 200,
    dateFilter: 'all',
    orderBy: 'visit_date_asc',
  });
  const { data: repairData } = useRepairSchedule(monthStart, monthEnd);

  const dateMap = useMemo(() => {
    const map = new Map<string, CalendarEvent[]>();
    for (const c of (storeData?.consultations || [])) {
      if (!c.visit_date) continue;
      if (c.visit_date < monthStart || c.visit_date > monthEnd) continue;
      if (!map.has(c.visit_date)) map.set(c.visit_date, []);
      map.get(c.visit_date)!.push({ kind: 'consult', data: c });
    }
    for (const c of (fieldData?.consultations || [])) {
      if (!c.visit_date) continue;
      if (c.visit_date < monthStart || c.visit_date > monthEnd) continue;
      if (!map.has(c.visit_date)) map.set(c.visit_date, []);
      map.get(c.visit_date)!.push({ kind: 'consult', data: c });
    }
    for (const r of (repairData || [])) {
      if (!r.visit_date) continue;
      if (!map.has(r.visit_date)) map.set(r.visit_date, []);
      map.get(r.visit_date)!.push({ kind: 'repair', data: r });
    }
    for (const arr of map.values()) {
      arr.sort((a, b) => {
        const tA = a.kind === 'consult' ? (a.data.visit_time || '') : (a.data.visit_time || '');
        const tB = b.kind === 'consult' ? (b.data.visit_time || '') : (b.data.visit_time || '');
        return tA.localeCompare(tB);
      });
    }
    return map;
  }, [storeData, fieldData, repairData, monthStart, monthEnd]);

  const calendarDays = getCalendarDays(year, month);
  const selectedEvents = dateMap.get(selectedDate) || [];

  const goMonth = (delta: number) => {
    const d = new Date(year, month + delta, 1);
    setYear(d.getFullYear());
    setMonth(d.getMonth());
    // 이동한 달의 1일을 선택일로 (오늘이 그 달이면 오늘)
    const nowY = today.getFullYear();
    const nowM = today.getMonth();
    if (d.getFullYear() === nowY && d.getMonth() === nowM) {
      setSelectedDate(todayStr);
    } else {
      setSelectedDate(formatYYYYMMDD(d.getFullYear(), d.getMonth(), 1));
    }
  };

  // 선택일 날짜 라벨 포맷
  const selectedLabel = (() => {
    const [y, m, d] = selectedDate.split('-').map(Number);
    const dt = new Date(y, m - 1, d);
    const weekday = DAY_NAMES[dt.getDay()];
    return { month: m, day: d, weekday };
  })();
  const isSelectedToday = selectedDate === todayStr;

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
      {/* 좌측: 달력 */}
      <div className="bg-white rounded-2xl border border-stone-200 p-4">
        <div className="flex items-center justify-between mb-3">
          <span className="text-sm font-bold text-stone-900">
            {year}년 {month + 1}월
          </span>
          <div className="flex items-center gap-1">
            <button
              onClick={() => goMonth(-1)}
              className="w-7 h-7 rounded-lg hover:bg-stone-100 flex items-center justify-center transition"
              aria-label="이전 달"
            >
              <ChevronLeft size={14} className="text-stone-500" />
            </button>
            <button
              onClick={() => goMonth(1)}
              className="w-7 h-7 rounded-lg hover:bg-stone-100 flex items-center justify-center transition"
              aria-label="다음 달"
            >
              <ChevronRight size={14} className="text-stone-500" />
            </button>
          </div>
        </div>

        <div className="flex items-center gap-3 mb-2 text-[10px] text-stone-500">
          <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />매장</span>
          <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-violet-500" />출장</span>
          <span className="flex items-center gap-1"><span className="w-1.5 h-1.5 rounded-full bg-amber-500" />수리</span>
        </div>

        <div className="grid grid-cols-7 gap-1 mb-1">
          {DAY_NAMES.map((d, i) => (
            <div key={d} className={`text-center text-[10px] font-semibold py-1 ${i === 0 ? 'text-rose-500' : i === 6 ? 'text-blue-500' : 'text-stone-500'}`}>{d}</div>
          ))}
        </div>

        <div className="grid grid-cols-7 gap-1">
          {calendarDays.map((day, idx) => {
            if (day === null) return <div key={`e-${idx}`} className="h-10" />;
            const dateStr = formatYYYYMMDD(year, month, day);
            const events = dateMap.get(dateStr) || [];
            const storeCount = events.filter((e) => e.kind === 'consult' && e.data.consultation_type === 'store_visit').length;
            const fieldCount = events.filter((e) => e.kind === 'consult' && e.data.consultation_type === 'field_request').length;
            const repairCount = events.filter((e) => e.kind === 'repair').length;
            const isToday = dateStr === todayStr;
            const isSelected = dateStr === selectedDate;
            return (
              <button
                key={dateStr}
                onClick={() => setSelectedDate(dateStr)}
                className={`h-10 rounded-lg p-1 text-[11px] transition flex flex-col items-center justify-start ${
                  isSelected ? 'bg-stone-900 text-white' :
                  isToday ? 'bg-amber-50 text-amber-900 border border-amber-300' :
                  'bg-white hover:bg-stone-50 text-stone-700'
                }`}
              >
                <span className="font-semibold leading-none mt-0.5">{day}</span>
                {events.length > 0 && (
                  <div className="flex items-center gap-0.5 mt-1">
                    {storeCount > 0 && <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-emerald-300' : 'bg-emerald-500'}`} />}
                    {fieldCount > 0 && <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-violet-300' : 'bg-violet-500'}`} />}
                    {repairCount > 0 && <span className={`w-1 h-1 rounded-full ${isSelected ? 'bg-amber-300' : 'bg-amber-500'}`} />}
                  </div>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* 우측: 선택일 타임라인 */}
      <div className="bg-white rounded-2xl border border-stone-200 p-4">
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <span className="text-sm font-bold text-stone-900">
              {selectedLabel.month}월 {selectedLabel.day}일
              <span className="text-stone-400 text-xs font-normal ml-1.5">({selectedLabel.weekday})</span>
            </span>
            {isSelectedToday && (
              <span className="text-[10px] px-2 py-0.5 rounded-full bg-amber-100 text-amber-700 font-semibold">오늘</span>
            )}
          </div>
          <span className="text-[10px] text-stone-500">{selectedEvents.length}건</span>
        </div>

        {selectedEvents.length === 0 ? (
          <div className="py-4 text-center text-xs text-stone-400">
            일정이 없습니다
          </div>
        ) : (
          <div className="relative pl-4">
            <div className="absolute left-1 top-1.5 bottom-1.5 w-px bg-stone-200" />
            <div className="space-y-2.5">
              {selectedEvents.map((ev) => {
                if (ev.kind === 'repair') {
                  const r = ev.data;
                  const qty = (r.qty_mamoru || 0) + (r.qty_other || 0);
                  return (
                    <div
                      key={`repair-${r.id}`}
                      className="relative flex items-center gap-2.5 group cursor-pointer"
                      onClick={() => router.push(`/repairs/${r.id}`)}
                    >
                      <div className="absolute -left-[15px] w-2.5 h-2.5 rounded-full bg-amber-500 ring-2 ring-white" />
                      <span className="text-xs font-semibold text-stone-500 w-10 shrink-0">{r.visit_time || '-'}</span>
                      <span className="text-[10px] px-1.5 py-0.5 rounded bg-amber-50 text-amber-700 font-semibold shrink-0">수리</span>
                      <div className="flex-1 min-w-0 flex items-center justify-between gap-2">
                        <span className="text-xs font-medium text-stone-800 truncate">{r.name}</span>
                        <span className="text-[10px] text-stone-400 truncate">
                          {qty > 0 ? `${qty}자루` : formatPhone(r.phone)}
                        </span>
                      </div>
                    </div>
                  );
                }
                const c = ev.data;
                const isStore = c.consultation_type === 'store_visit';
                return (
                  <div
                    key={`consult-${c.id}`}
                    className="relative flex items-center gap-2.5 group cursor-pointer"
                    onClick={() => setDetail(c)}
                  >
                    <div className={`absolute -left-[15px] w-2.5 h-2.5 rounded-full ring-2 ring-white ${isStore ? 'bg-emerald-500' : 'bg-violet-500'}`} />
                    <span className="text-xs font-semibold text-stone-500 w-10 shrink-0">{c.visit_time || '-'}</span>
                    <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold shrink-0 ${isStore ? 'bg-emerald-50 text-emerald-700' : 'bg-violet-50 text-violet-700'}`}>
                      {isStore ? '매장' : '출장'}
                    </span>
                    <div className="flex-1 min-w-0 flex items-center justify-between gap-2">
                      <span className="text-xs font-medium text-stone-800 truncate">{activityDisplay(c.activity_name, c.name)}</span>
                      <span className="text-[10px] text-stone-400 truncate">{formatPhone(c.phone || '')}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
      </div>

      {/* ── 일정 상세 모달 ── */}
      {detail && (() => {
        const isStore = detail.consultation_type === 'store_visit';
        const addr = [detail.address_road, detail.address_detail].filter(Boolean).join(' ');
        return (
          <div className="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4" onClick={() => setDetail(null)}>
            <div className="bg-white rounded-2xl w-full max-w-md max-h-[85vh] overflow-y-auto shadow-xl" onClick={(e) => e.stopPropagation()}>
              {/* 헤더 */}
              <div className="p-4 border-b border-stone-100 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 flex-wrap mb-1">
                    <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold ${isStore ? 'bg-emerald-50 text-emerald-700' : 'bg-violet-50 text-violet-700'}`}>{isStore ? '매장방문' : '출장'}</span>
                    <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold ${CONSULTATION_STATUS_COLOR[detail.status] || 'bg-stone-100 text-stone-600'}`}>{CONSULTATION_STATUS_LABEL[detail.status] || detail.status}</span>
                  </div>
                  <h3 className="text-base font-bold text-stone-900 truncate">{activityDisplay(detail.activity_name, detail.name)}</h3>
                  {(ctx?.company || detail.position) && (
                    <p className="text-xs text-stone-500 mt-0.5 truncate">{[ctx?.company, detail.position].filter(Boolean).join(' · ')}</p>
                  )}
                </div>
                <button onClick={() => setDetail(null)} className="text-stone-400 hover:text-stone-700 shrink-0" aria-label="닫기"><X size={18} /></button>
              </div>

              {/* 기본 정보 */}
              <div className="p-4 space-y-3">
                <div className="flex items-center gap-2 text-sm">
                  <CalendarIcon size={15} className="text-stone-400 shrink-0" />
                  <span className="text-stone-500 w-12 shrink-0 text-xs">일시</span>
                  <span className="font-medium text-stone-800">{detail.visit_date || '-'}{detail.visit_time ? ` ${detail.visit_time}` : ''}</span>
                </div>
                <div className="flex items-center gap-2 text-sm">
                  <Phone size={15} className="text-stone-400 shrink-0" />
                  <span className="text-stone-500 w-12 shrink-0 text-xs">연락처</span>
                  {detail.phone
                    ? <a href={`tel:${detail.phone}`} className="font-medium text-blue-600 hover:underline">{formatPhone(detail.phone)}</a>
                    : <span className="text-stone-400">-</span>}
                </div>
                {addr && (
                  <div className="flex items-start gap-2 text-sm">
                    <MapPin size={15} className="text-stone-400 shrink-0 mt-0.5" />
                    <span className="text-stone-500 w-12 shrink-0 text-xs mt-0.5">주소</span>
                    <span className="font-medium text-stone-800">{addr}</span>
                  </div>
                )}

                {/* 고객 메모 */}
                {detail.memo && (
                  <div className="pt-1">
                    <p className="text-[11px] font-semibold text-stone-500 mb-1 flex items-center gap-1"><StickyNote size={12} />고객 메모</p>
                    <p className="text-sm text-stone-700 whitespace-pre-wrap bg-stone-50 rounded-lg p-3 leading-relaxed">{detail.memo}</p>
                  </div>
                )}
                {/* 관리자(내) 메모 */}
                {detail.admin_note && (
                  <div className="pt-1">
                    <p className="text-[11px] font-semibold text-amber-600 mb-1 flex items-center gap-1"><StickyNote size={12} />내 메모 (관리자 전용)</p>
                    <p className="text-sm text-stone-700 whitespace-pre-wrap bg-amber-50 rounded-lg p-3 leading-relaxed">{detail.admin_note}</p>
                  </div>
                )}
                {!detail.memo && !detail.admin_note && (
                  <p className="text-xs text-stone-400 pt-1">작성된 메모가 없습니다.</p>
                )}
              </div>

              {/* 고객 상담 메모 (특징·불편·요구 등 — 빠른 입력) */}
              {detail.customer_id && (
                <div className="px-4 pb-4 border-t border-stone-100 pt-3">
                  <p className="text-[11px] font-semibold text-stone-500 mb-2">고객 상담 메모</p>
                  <CustomerNotes customerId={detail.customer_id} collapsed />
                </div>
              )}

              {/* 이력 (구매 · 이전 상담) — 누르면 상세로 이동 */}
              {(ctx && (ctx.sales.length > 0 || ctx.pastConsults.length > 0)) ? (
                <div className="px-4 pb-4 space-y-3 border-t border-stone-100 pt-3">
                  {ctx.sales.length > 0 && (
                    <div>
                      <p className="text-[11px] font-semibold text-stone-500 mb-1.5">구매 이력 {ctx.sales.length}건</p>
                      <div className="space-y-1">
                        {ctx.sales.map((s) => (
                          <button key={s.id} onClick={() => router.push(`/sales/${s.id}`)}
                            className="w-full flex items-center justify-between gap-2 px-2.5 py-2 rounded-lg border border-stone-100 hover:bg-stone-50 hover:border-stone-200 text-left transition">
                            <span className="text-xs text-stone-600 truncate">{s.sale_date} · {s.sale_number}</span>
                            <span className="text-xs font-semibold text-stone-800 shrink-0">₩{(s.total_amount || 0).toLocaleString()}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                  {ctx.pastConsults.length > 0 && (
                    <div>
                      <p className="text-[11px] font-semibold text-stone-500 mb-1.5">이전 상담 {ctx.pastConsults.length}건</p>
                      <div className="space-y-1">
                        {ctx.pastConsults.map((pc) => (
                          <button key={pc.id} onClick={() => router.push(`/consultations/${pc.id}`)}
                            className="w-full flex items-center justify-between gap-2 px-2.5 py-2 rounded-lg border border-stone-100 hover:bg-stone-50 hover:border-stone-200 text-left transition">
                            <span className="text-xs text-stone-600 truncate">{pc.visit_date || '-'}{pc.visit_time ? ` ${pc.visit_time}` : ''} · {CTYPE_LABEL[pc.consultation_type] || pc.consultation_type}</span>
                            <span className={`text-[10px] px-1.5 py-0.5 rounded shrink-0 ${CONSULTATION_STATUS_COLOR[pc.status] || 'bg-stone-100 text-stone-600'}`}>{CONSULTATION_STATUS_LABEL[pc.status] || pc.status}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              ) : ctxLoading ? (
                <div className="px-4 pb-4 text-xs text-stone-400 border-t border-stone-100 pt-3">이력 불러오는 중…</div>
              ) : null}

              {/* 푸터 */}
              <div className="p-3 border-t border-stone-100 flex items-center gap-2 sticky bottom-0 bg-white">
                <button onClick={() => setDetail(null)} className="flex-1 px-3 py-2 rounded-lg border border-stone-200 text-sm text-stone-600 hover:bg-stone-50">닫기</button>
                <button onClick={() => router.push(`/consultations/${detail.id}`)} className="flex-1 px-3 py-2 rounded-lg bg-stone-900 text-white text-sm font-semibold hover:bg-stone-800 flex items-center justify-center gap-1">상세 페이지 <ArrowRight size={14} /></button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}
