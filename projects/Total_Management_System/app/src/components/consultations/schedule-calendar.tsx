'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { activityDisplay } from '@/lib/customer/display';
import { getCalendarDays, ymd } from '@/lib/schedule/calendar';
import { useScheduleEvents } from '@/lib/schedule/use-schedule-events';
import { SCHEDULE_COLORS } from '@/lib/schedule/colors';
import { formatPhone } from '@/lib/utils/format';
import { ChevronLeft, ChevronRight, Calendar, Store, Truck, Wrench, Package } from 'lucide-react';

// 2026-05-25 Phase 3-A: 달력 일정 통합 (매장방문 / 출장 / 직접방문)
// 2026-08-12: 공용 useScheduleEvents로 통합(DRY) + 방문수거(수거예정) 추가
const DAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

interface ScheduleCalendarProps {
  onSelect?: (id: string) => void;
}

export function ScheduleCalendar({ onSelect }: ScheduleCalendarProps = {}) {
  const router = useRouter();
  const today = new Date();
  const [year, setYear] = useState(today.getFullYear());
  const [month, setMonth] = useState(today.getMonth());
  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  const todayStr = ymd(today.getFullYear(), today.getMonth(), today.getDate());

  // 해당 월 범위 계산
  const monthStart = ymd(year, month, 1);
  const monthEnd = ymd(year, month, new Date(year, month + 1, 0).getDate());

  // 상담(매장·출장) + 복원수리(직접방문·방문수거) 통합 조회
  const { dateMap } = useScheduleEvents(monthStart, monthEnd);

  const calendarDays = getCalendarDays(year, month);

  // 선택된 날짜의 일정 목록
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

          const dateStr = ymd(year, month, day);
          const events = dateMap.get(dateStr) || [];
          const isToday = dateStr === todayStr;
          const isSelected = dateStr === selectedDate;
          const storeCount = events.filter((e) => e.kind === 'consult' && e.data.consultation_type === 'store_visit').length;
          const fieldCount = events.filter((e) => e.kind === 'consult' && e.data.consultation_type === 'field_request').length;
          const repairCount = events.filter((e) => e.kind === 'repair').length;
          const pickupCount = events.filter((e) => e.kind === 'pickup').length;
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
                    <span className={`w-1.5 h-1.5 rounded-full ${SCHEDULE_COLORS.store.dot}`} title={`매장방문 ${storeCount}건`} />
                  )}
                  {fieldCount > 0 && (
                    <span className={`w-1.5 h-1.5 rounded-full ${SCHEDULE_COLORS.field.dot}`} title={`출장요청 ${fieldCount}건`} />
                  )}
                  {repairCount > 0 && (
                    <span className={`w-1.5 h-1.5 rounded-full ${SCHEDULE_COLORS.repair_visit.dot}`} title={`복원수리 직접방문 ${repairCount}건`} />
                  )}
                  {pickupCount > 0 && (
                    <span className={`w-1.5 h-1.5 rounded-full ${SCHEDULE_COLORS.repair_pickup.dot}`} title={`복원수리 수거예정 ${pickupCount}건`} />
                  )}
                </div>
              )}
            </button>
          );
        })}
      </div>

      {/* 범례 — 4종 통합 (수거예정 추가 2026-08-12) */}
      <div className="flex items-center gap-3 mt-3 text-xs text-neutral-500 flex-wrap">
        <span className="flex items-center gap-1">
          <span className={`w-2 h-2 rounded-full ${SCHEDULE_COLORS.store.dot}`} />
          매장방문
        </span>
        <span className="flex items-center gap-1">
          <span className={`w-2 h-2 rounded-full ${SCHEDULE_COLORS.field.dot}`} />
          출장요청
        </span>
        <span className="flex items-center gap-1">
          <span className={`w-2 h-2 rounded-full ${SCHEDULE_COLORS.repair_visit.dot}`} />
          복원수리 직접방문
        </span>
        <span className="flex items-center gap-1">
          <span className={`w-2 h-2 rounded-full ${SCHEDULE_COLORS.repair_pickup.dot}`} />
          수거예정
        </span>
      </div>

      {/* 선택된 날짜 일정 목록 — 4종 분기 */}
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
                      <Wrench size={14} className={`${SCHEDULE_COLORS.repair_visit.text} shrink-0`} />
                      <div className="flex-1 min-w-0">
                        <span className="text-sm font-semibold text-indigo-black">{r.name}</span>
                        <span className="text-xs text-neutral-400 ml-2">
                          {r.visit_time || ''}
                          {r.visit_duration_min ? ` · ${r.visit_duration_min}분` : ''}
                          {qty > 0 ? ` · ${qty}자루` : ''}
                        </span>
                      </div>
                      <span className="text-xs text-neutral-500">{formatPhone(r.phone)}</span>
                      <Badge className={SCHEDULE_COLORS.repair_visit.badge}>직접방문</Badge>
                    </div>
                  );
                }
                if (ev.kind === 'pickup') {
                  const p = ev.data;
                  const qty = (p.qty_mamoru || 0) + (p.qty_other || 0);
                  return (
                    <div
                      key={`pickup-${p.id}`}
                      className="flex items-center gap-2 p-2 rounded-lg hover:bg-warm-ivory/60 cursor-pointer transition"
                      onClick={() => router.push(`/repairs/${p.id}`)}
                    >
                      <Package size={14} className={`${SCHEDULE_COLORS.repair_pickup.text} shrink-0`} />
                      <div className="flex-1 min-w-0">
                        <span className="text-sm font-semibold text-indigo-black">{p.name}</span>
                        <span className="text-xs text-neutral-400 ml-2">
                          {qty > 0 ? `${qty}자루` : ''}
                        </span>
                      </div>
                      <span className="text-xs text-neutral-500">{formatPhone(p.phone)}</span>
                      <Badge className={SCHEDULE_COLORS.repair_pickup.badge}>수거</Badge>
                    </div>
                  );
                }
                if (ev.kind === 'return_pickup') {
                  const p = ev.data;
                  return (
                    <div
                      key={`return-${p.id}`}
                      className="flex items-center gap-2 p-2 rounded-lg hover:bg-warm-ivory/60 cursor-pointer transition"
                      onClick={() => router.push(`/returns/${p.id}`)}
                    >
                      <Package size={14} className={`${SCHEDULE_COLORS.return_pickup.text} shrink-0`} />
                      <div className="flex-1 min-w-0">
                        <span className="text-sm font-semibold text-indigo-black">{p.name || '고객'}</span>
                        <span className="text-xs text-neutral-400 ml-2">{p.pickup_method || ''}</span>
                      </div>
                      <span className="text-xs text-neutral-500">{formatPhone(p.phone || '')}</span>
                      <Badge className={SCHEDULE_COLORS.return_pickup.badge}>반품수거</Badge>
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
                      <Store size={14} className={`${SCHEDULE_COLORS.store.text} shrink-0`} />
                    ) : (
                      <Truck size={14} className={`${SCHEDULE_COLORS.field.text} shrink-0`} />
                    )}
                    <div className="flex-1 min-w-0">
                      <span className="text-sm font-semibold text-indigo-black">{activityDisplay(c.activity_name, c.name)}</span>
                      <span className="text-xs text-neutral-400 ml-2">{c.visit_time || ''}</span>
                    </div>
                    <span className="text-xs text-neutral-500">{formatPhone(c.phone)}</span>
                    <Badge className={c.consultation_type === 'store_visit' ? SCHEDULE_COLORS.store.badge : SCHEDULE_COLORS.field.badge}>
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
