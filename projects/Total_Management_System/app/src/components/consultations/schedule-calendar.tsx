'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { useConsultations } from '@/hooks/use-consultations';
import { useRepairSchedule, type RepairScheduleItem } from '@/hooks/use-repairs';
import { formatPhone } from '@/lib/utils/format';
import { ChevronLeft, ChevronRight, Calendar, Store, Truck, Wrench } from 'lucide-react';
import type { Consultation } from '@/lib/supabase/types';

// 2026-05-25 Phase 3-A: 달력 일정 통합 (매장방문 / 출장 / 직접방문)
type CalendarEvent =
  | { kind: 'consult'; data: Consultation }
  | { kind: 'repair';  data: RepairScheduleItem };

const DAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

/** 해당 월의 달력 그리드 생성 */
function getCalendarDays(year: number, month: number) {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const startOffset = firstDay.getDay(); // 시작 요일 (0=일)
  const totalDays = lastDay.getDate();

  const days: (number | null)[] = [];
  for (let i = 0; i < startOffset; i++) days.push(null);
  for (let d = 1; d <= totalDays; d++) days.push(d);
  // 끝 채우기 (7의 배수)
  while (days.length % 7 !== 0) days.push(null);
  return days;
}

function formatYYYYMMDD(year: number, month: number, day: number): string {
  return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

interface ScheduleCalendarProps {
  onSelect?: (id: string) => void;
}

export function ScheduleCalendar({ onSelect }: ScheduleCalendarProps = {}) {
  const router = useRouter();
  const today = new Date();
  const [year, setYear] = useState(today.getFullYear());
  const [month, setMonth] = useState(today.getMonth());
  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  const todayStr = formatYYYYMMDD(today.getFullYear(), today.getMonth(), today.getDate());

  // 해당 월 범위 계산
  const monthStart = formatYYYYMMDD(year, month, 1);
  const monthEnd = formatYYYYMMDD(year, month, new Date(year, month + 1, 0).getDate());

  // confirmed + suggested 매장방문 + 출장요청 조회
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
  // 2026-05-25 Phase 3-A: 복원수리 직접방문 일정도 함께 표시
  const { data: repairData } = useRepairSchedule(monthStart, monthEnd);

  // 날짜별 일정 그룹핑 (3종 통합)
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
      // 위 hook 이 monthStart/monthEnd 로 필터링하므로 추가 범위 검사 불필요
      if (!map.has(r.visit_date)) map.set(r.visit_date, []);
      map.get(r.visit_date)!.push({ kind: 'repair', data: r });
    }
    // 시간 순 정렬 (각 날짜 안)
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

  // 선택된 날짜의 상담 목록
  const selectedConsultations = selectedDate ? dateMap.get(selectedDate) || [] : [];

  const goMonth = (delta: number) => {
    const d = new Date(year, month + delta, 1);
    setYear(d.getFullYear());
    setMonth(d.getMonth());
    setSelectedDate(null);
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle>
          <Calendar size={16} className="inline mr-1.5" />
          확정 일정
        </CardTitle>
      </CardHeader>

      {/* 월 네비게이션 */}
      <div className="flex items-center justify-between mb-3">
        <Button variant="ghost" size="sm" onClick={() => goMonth(-1)}>
          <ChevronLeft size={16} />
        </Button>
        <span className="text-sm font-bold text-indigo-black">
          {year}년 {month + 1}월
        </span>
        <Button variant="ghost" size="sm" onClick={() => goMonth(1)}>
          <ChevronRight size={16} />
        </Button>
      </div>

      {/* 요일 헤더 */}
      <div className="grid grid-cols-7 text-center text-xs font-semibold text-neutral-400 mb-1">
        {DAY_NAMES.map((d) => (
          <div key={d} className={d === '일' ? 'text-error' : d === '토' ? 'text-info' : ''}>
            {d}
          </div>
        ))}
      </div>

      {/* 달력 그리드 */}
      <div className="grid grid-cols-7 gap-px">
        {calendarDays.map((day, idx) => {
          if (day === null) {
            return <div key={`e-${idx}`} className="h-12" />;
          }

          const dateStr = formatYYYYMMDD(year, month, day);
          const events = dateMap.get(dateStr) || [];
          const isToday = dateStr === todayStr;
          const isSelected = dateStr === selectedDate;
          const storeCount = events.filter((e) => e.kind === 'consult' && e.data.consultation_type === 'store_visit').length;
          const fieldCount = events.filter((e) => e.kind === 'consult' && e.data.consultation_type === 'field_request').length;
          const repairCount = events.filter((e) => e.kind === 'repair').length;
          const dayOfWeek = (idx % 7);

          return (
            <button
              key={dateStr}
              onClick={() => setSelectedDate(isSelected ? null : dateStr)}
              className={`h-12 flex flex-col items-center justify-start pt-1 rounded-md text-xs transition hover:bg-warm-ivory/80 ${
                isSelected ? 'bg-terracotta/10 ring-1 ring-terracotta' : ''
              } ${isToday ? 'bg-terracotta/5' : ''}`}
            >
              <span className={`font-semibold ${
                isToday ? 'text-terracotta'
                : dayOfWeek === 0 ? 'text-error'
                : dayOfWeek === 6 ? 'text-info'
                : 'text-neutral-700'
              }`}>
                {day}
              </span>
              {events.length > 0 && (
                <div className="flex gap-0.5 mt-0.5">
                  {storeCount > 0 && (
                    <span className="w-1.5 h-1.5 rounded-full bg-green-500" title={`매장방문 ${storeCount}건`} />
                  )}
                  {fieldCount > 0 && (
                    <span className="w-1.5 h-1.5 rounded-full bg-purple-500" title={`출장요청 ${fieldCount}건`} />
                  )}
                  {repairCount > 0 && (
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-500" title={`복원수리 직접방문 ${repairCount}건`} />
                  )}
                </div>
              )}
            </button>
          );
        })}
      </div>

      {/* 범례 — 3종 통합 (Phase 3-A, 2026-05-25) */}
      <div className="flex items-center gap-3 mt-3 text-xs text-neutral-500 flex-wrap">
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-green-500" />
          매장방문
        </span>
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-purple-500" />
          출장요청
        </span>
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded-full bg-amber-500" />
          복원수리 직접방문
        </span>
      </div>

      {/* 선택된 날짜 일정 목록 — 3종 분기 (Phase 3-A) */}
      {selectedDate && (
        <div className="mt-4 border-t border-neutral-100 pt-3">
          <h4 className="text-xs font-bold text-neutral-600 mb-2">
            {selectedDate} 일정 ({selectedConsultations.length}건)
          </h4>
          {selectedConsultations.length === 0 ? (
            <p className="text-xs text-neutral-400">일정이 없습니다</p>
          ) : (
            <div className="space-y-2">
              {selectedConsultations.map((ev) => {
                if (ev.kind === 'repair') {
                  const r = ev.data;
                  const qty = (r.qty_mamoru || 0) + (r.qty_other || 0);
                  return (
                    <div
                      key={`repair-${r.id}`}
                      className="flex items-center gap-2 p-2 rounded-lg hover:bg-warm-ivory/60 cursor-pointer transition"
                      onClick={() => router.push(`/repairs/${r.id}`)}
                    >
                      <Wrench size={14} className="text-amber-600 shrink-0" />
                      <div className="flex-1 min-w-0">
                        <span className="text-sm font-semibold text-indigo-black">{r.name}</span>
                        <span className="text-xs text-neutral-400 ml-2">
                          {r.visit_time || ''}
                          {r.visit_duration_min ? ` · ${r.visit_duration_min}분` : ''}
                          {qty > 0 ? ` · ${qty}자루` : ''}
                        </span>
                      </div>
                      <span className="text-xs text-neutral-500">{formatPhone(r.phone)}</span>
                      <Badge className="bg-amber-100 text-amber-700">직접방문</Badge>
                    </div>
                  );
                }
                const c = ev.data;
                return (
                  <div
                    key={`consult-${c.id}`}
                    className="flex items-center gap-2 p-2 rounded-lg hover:bg-warm-ivory/60 cursor-pointer transition"
                    onClick={() => onSelect ? onSelect(c.id) : router.push(`/consultations/${c.id}`)}
                  >
                    {c.consultation_type === 'store_visit' ? (
                      <Store size={14} className="text-green-600 shrink-0" />
                    ) : (
                      <Truck size={14} className="text-purple-600 shrink-0" />
                    )}
                    <div className="flex-1 min-w-0">
                      <span className="text-sm font-semibold text-indigo-black">{c.name}</span>
                      <span className="text-xs text-neutral-400 ml-2">{c.visit_time || ''}</span>
                    </div>
                    <span className="text-xs text-neutral-500">{formatPhone(c.phone)}</span>
                    <Badge className={c.consultation_type === 'store_visit' ? 'bg-green-100 text-green-700' : 'bg-purple-100 text-purple-700'}>
                      {c.consultation_type === 'store_visit' ? '매장' : '출장'}
                    </Badge>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </Card>
  );
}
