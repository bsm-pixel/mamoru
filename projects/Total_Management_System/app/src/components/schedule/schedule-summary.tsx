'use client';

/**
 * ScheduleSummary — 일정 페이지 상단 요약 스트립 (2026-08-12)
 *   오늘/이번주 건수 + 종류별(이번주) + '다음 일정' D-day.
 *   /schedule 전용(대시보드엔 별도 KPI가 있어 미표시). 공용 useScheduleEvents 재사용.
 */

import Link from 'next/link';
import { useScheduleEvents, flattenEvents, eventCategory, eventName, eventTimeStr, eventHref } from '@/lib/schedule/use-schedule-events';
import { SCHEDULE_COLORS, type ScheduleCategory } from '@/lib/schedule/colors';
import { ymd, parseYmd } from '@/lib/schedule/calendar';
import { CalendarClock, ArrowRight } from 'lucide-react';

const CATS: ScheduleCategory[] = ['store', 'field', 'repair_visit', 'repair_pickup'];

export function ScheduleSummary() {
  const now = new Date();
  const todayStr = ymd(now.getFullYear(), now.getMonth(), now.getDate());
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 45);
  const endStr = ymd(end.getFullYear(), end.getMonth(), end.getDate());
  const dow = now.getDay();
  const ws = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow);
  const we = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dow + 6);
  const weekStart = ymd(ws.getFullYear(), ws.getMonth(), ws.getDate());
  const weekEnd = ymd(we.getFullYear(), we.getMonth(), we.getDate());
  const nowHHMM = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  const { dateMap } = useScheduleEvents(todayStr, endStr);
  const flat = flattenEvents(dateMap);

  const todayCount = flat.filter((x) => x.date === todayStr).length;
  const weekItems = flat.filter((x) => x.date >= weekStart && x.date <= weekEnd);

  const catCount: Record<ScheduleCategory, number> = { store: 0, field: 0, repair_visit: 0, repair_pickup: 0 };
  for (const x of weekItems) catCount[eventCategory(x.ev)]++;

  // 다음 일정 (오늘 이미 지난 시간은 제외)
  const next = flat.find((x) => x.date > todayStr || (x.date === todayStr && (eventTimeStr(x.ev) ?? '99:99') >= nowHHMM)) ?? null;

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4">
      <div className="flex flex-wrap items-center gap-x-5 gap-y-3">
        {/* 오늘 / 이번주 */}
        <div className="flex items-center gap-4">
          <div>
            <div className="text-[11px] text-stone-400">오늘</div>
            <div className="text-xl font-bold text-stone-900 leading-none mt-0.5">{todayCount}<span className="text-xs font-normal text-stone-400 ml-0.5">건</span></div>
          </div>
          <div className="w-px h-8 bg-stone-200" />
          <div>
            <div className="text-[11px] text-stone-400">이번주</div>
            <div className="text-xl font-bold text-stone-900 leading-none mt-0.5">{weekItems.length}<span className="text-xs font-normal text-stone-400 ml-0.5">건</span></div>
          </div>
        </div>

        {/* 종류별(이번주) */}
        <div className="flex items-center gap-2.5 flex-wrap">
          {CATS.filter((c) => catCount[c] > 0).map((c) => (
            <span key={c} className="flex items-center gap-1 text-xs text-stone-500">
              <span className={`w-1.5 h-1.5 rounded-full ${SCHEDULE_COLORS[c].dot}`} />
              {SCHEDULE_COLORS[c].label} {catCount[c]}
            </span>
          ))}
          {weekItems.length === 0 && <span className="text-xs text-stone-300">이번주 일정 없음</span>}
        </div>

        {/* 다음 일정 */}
        <div className="ml-auto">
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
            <span className="text-xs text-stone-400">예정된 일정 없음</span>
          )}
        </div>
      </div>
    </div>
  );
}
