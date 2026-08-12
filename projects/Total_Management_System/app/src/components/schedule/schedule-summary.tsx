'use client';

/**
 * ScheduleSummary — 일정 페이지 상단 요약 (2026-08-12, '오늘 집중형' 개편)
 *   - 오늘: 실제 일정(종류+시간+이름) 미니 리스트 (숫자만 표기하던 모호함 해소)
 *   - 이번주: 종류별(매장·출장·수리·수거) 칩 — 0도 흐리게 항상 표시
 *   - 다음 일정: 오늘 이후 가장 가까운 일정 pill
 *   /schedule 전용(대시보드엔 별도 KPI). 공용 useScheduleEvents 재사용.
 */

import Link from 'next/link';
import { useScheduleEvents, flattenEvents, eventCategory, eventName, eventTimeStr, eventHref } from '@/lib/schedule/use-schedule-events';
import { SCHEDULE_COLORS, type ScheduleCategory } from '@/lib/schedule/colors';
import { ymd, parseYmd } from '@/lib/schedule/calendar';
import { CalendarClock, ArrowRight } from 'lucide-react';

const CATS: ScheduleCategory[] = ['store', 'field', 'repair_visit', 'repair_pickup'];
const TODAY_MAX = 3;

export function ScheduleSummary() {
  const now = new Date();
  const todayStr = ymd(now.getFullYear(), now.getMonth(), now.getDate());
  const dow = now.getDay();
  const ws = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow);
  const we = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow + 6);
  const weekStart = ymd(ws.getFullYear(), ws.getMonth(), ws.getDate());
  const weekEnd = ymd(we.getFullYear(), we.getMonth(), we.getDate());
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 45);
  const endStr = ymd(end.getFullYear(), end.getMonth(), end.getDate());

  const { dateMap } = useScheduleEvents(weekStart, endStr);
  const flat = flattenEvents(dateMap);

  const todayItems = flat.filter((x) => x.date === todayStr);
  const weekItems = flat.filter((x) => x.date >= weekStart && x.date <= weekEnd);

  const catCount: Record<ScheduleCategory, number> = { store: 0, field: 0, repair_visit: 0, repair_pickup: 0 };
  for (const x of weekItems) catCount[eventCategory(x.ev)]++;

  // 다음 일정 = 오늘 이후(내일부터) 가장 가까운 건 (오늘은 위 리스트에 이미 표시)
  const next = flat.find((x) => x.date > todayStr) ?? null;

  const tp = parseYmd(todayStr);
  const shownToday = todayItems.slice(0, TODAY_MAX);
  const extraToday = todayItems.length - shownToday.length;

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4">
      {/* 상단: 오늘 일정 + 다음 일정 */}
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2 mb-2">
            <span className="text-sm font-bold text-stone-900">오늘</span>
            <span className="text-xs text-stone-400">{tp.m}/{tp.d}({tp.dow})</span>
            {todayItems.length > 0 && (
              <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-stone-100 text-stone-500 font-semibold">{todayItems.length}건</span>
            )}
          </div>
          {todayItems.length === 0 ? (
            <p className="text-xs text-stone-400">오늘은 일정이 없습니다</p>
          ) : (
            <div className="space-y-1.5">
              {shownToday.map(({ ev }, i) => {
                const color = SCHEDULE_COLORS[eventCategory(ev)];
                const time = eventTimeStr(ev);
                return (
                  <Link key={i} href={eventHref(ev)} className="flex items-center gap-2 group">
                    <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${color.dot}`} />
                    <span className="text-xs font-semibold text-stone-500 w-10 shrink-0">{time || '-'}</span>
                    <span className={`text-[10px] px-1.5 py-0.5 rounded font-semibold shrink-0 ${color.badge}`}>{color.label}</span>
                    <span className="text-sm font-medium text-stone-800 truncate group-hover:underline">{eventName(ev)}</span>
                  </Link>
                );
              })}
              {extraToday > 0 && <p className="text-[11px] text-stone-400 pl-[22px]">외 {extraToday}건</p>}
            </div>
          )}
        </div>

        {/* 다음 일정 */}
        <div className="shrink-0">
          {next ? (
            <Link href={eventHref(next.ev)} className="flex items-center gap-2 bg-stone-50 hover:bg-stone-100 rounded-xl px-3 py-2 transition group">
              <CalendarClock size={15} className="text-stone-400 shrink-0" />
              <div className="min-w-0">
                <div className="text-[10px] text-stone-400 leading-none">다음 일정</div>
                <div className="text-xs font-semibold text-stone-800 truncate mt-0.5">
                  {(() => { const p = parseYmd(next.date); return `${p.m}/${p.d}(${p.dow})`; })()}
                  {eventTimeStr(next.ev) ? ` ${eventTimeStr(next.ev)}` : ''}
                  {' · '}{SCHEDULE_COLORS[eventCategory(next.ev)].label} {eventName(next.ev)}
                </div>
              </div>
              <ArrowRight size={14} className="text-stone-400 shrink-0 group-hover:translate-x-0.5 transition-transform" />
            </Link>
          ) : (
            <span className="text-xs text-stone-400 whitespace-nowrap">이후 예정 없음</span>
          )}
        </div>
      </div>

      {/* 이번주 종류별 */}
      <div className="mt-3 pt-3 border-t border-stone-100 flex items-center gap-x-3 gap-y-1.5 flex-wrap">
        <span className="text-[11px] font-semibold text-stone-400">이번주</span>
        {CATS.map((c) => {
          const has = catCount[c] > 0;
          return (
            <span key={c} className={`flex items-center gap-1 text-xs ${has ? 'text-stone-700 font-medium' : 'text-stone-300'}`}>
              <span className={`w-1.5 h-1.5 rounded-full ${has ? SCHEDULE_COLORS[c].dot : 'bg-stone-200'}`} />
              {SCHEDULE_COLORS[c].label} {catCount[c]}
            </span>
          );
        })}
        <span className="text-[11px] text-stone-400 ml-auto">총 {weekItems.length}건</span>
      </div>
    </div>
  );
}
