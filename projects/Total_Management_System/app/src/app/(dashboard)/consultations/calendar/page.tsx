'use client';

/**
 * 상담 달력 관리 (078)
 *
 * 사장님이 4개월 달력(현재월 ~ +3개월)을 보고 클릭 한 번으로 휴무일을 막거나 해제.
 * 막힌 날짜는 고객 셀프 예약 흐름(매장/출장/톡상담 폼)에서 자동 비활성화.
 *
 * 사장님 룰: 사장님 측 흐름(일정수동등록/시간제안)은 막힘 무시 — 항상 유동.
 *   상세: memory/feedback_consultation_blackout_rule.md
 */

import { useState, useMemo, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Modal } from '@/components/ui/modal';
import { CalendarOff, Store, Truck, Info, AlertTriangle, Clock, Ban } from 'lucide-react';
import {
  useBlackouts,
  useCreateBlackout,
  useDeleteBlackout,
  useConsultationSettings,
  useUpdateDisabledWeekdays,
  useUpdateBusinessHours,
  useBlockedSlots,
  useReplaceBlockedSlots,
  type BlackoutConsultation,
  type BlockedSlot,
} from '@/hooks/use-blackouts';
import { formatPhone } from '@/lib/utils/format';

const DAY_NAMES = ['일', '월', '화', '수', '목', '금', '토'];

// ── 118: 시간 격자(30분) 유틸 — 칠한 칸 ↔ 구간 상호 변환 ─────────────────
const SLOT = 30;
const toMin = (t: string) => { const [h, m] = t.split(':').map(Number); return h * 60 + m; };
const toHHMM = (m: number) => `${pad2(Math.floor(m / 60))}:${pad2(m % 60)}`;

/** 영업시간 범위의 30분 칸 시작분 목록 */
function gridCells(startHour: number, endHour: number): number[] {
  const out: number[] = [];
  for (let m = startHour * 60; m + SLOT <= endHour * 60; m += SLOT) out.push(m);
  return out;
}
/** 차단 구간들 → 칠해진 칸 Set */
function rangesToCells(ranges: { start_time: string; end_time: string }[]): Set<number> {
  const s = new Set<number>();
  for (const r of ranges) {
    for (let m = toMin(r.start_time); m < toMin(r.end_time); m += SLOT) s.add(m);
  }
  return s;
}
/** 칠해진 칸 Set → 연속 구간으로 병합 (저장용) */
function cellsToRanges(cells: Set<number>): { start_time: string; end_time: string }[] {
  const sorted = [...cells].sort((a, b) => a - b);
  const out: { start_time: string; end_time: string }[] = [];
  let i = 0;
  while (i < sorted.length) {
    const start = sorted[i];
    let end = start + SLOT;
    while (i + 1 < sorted.length && sorted[i + 1] === end) { end += SLOT; i++; }
    out.push({ start_time: toHHMM(start), end_time: toHHMM(end) });
    i++;
  }
  return out;
}
/** 차단 구간 요약 문자열 — 달력 셀/목록용 (예: "12:00~16:00 +1") */
function summarizeRanges(ranges: { start_time: string; end_time: string }[]): string {
  if (ranges.length === 0) return '';
  const sorted = [...ranges].sort((a, b) => a.start_time.localeCompare(b.start_time));
  const first = `${sorted[0].start_time}~${sorted[0].end_time}`;
  return sorted.length > 1 ? `${first} +${sorted.length - 1}` : first;
}

/** 118: 매주 반복 시간차단 카드 — 요일 탭 + 시간 격자 */
function WeeklyBlockCard({ startHour, endHour, weeklyMap, disabledWeekdays }: {
  startHour: number; endHour: number;
  weeklyMap: Map<number, BlockedSlot[]>;
  disabledWeekdays: number[];
}) {
  const [wd, setWd] = useState(1); // 기본 월요일
  const replace = useReplaceBlockedSlots();
  const saved = useMemo(() => weeklyMap.get(wd) || [], [weeklyMap, wd]);
  const [cells, setCells] = useState<Set<number>>(new Set());
  useEffect(() => { setCells(rangesToCells(saved)); }, [saved]);

  const savedCells = useMemo(() => rangesToCells(saved), [saved]);
  const dirty = cells.size !== savedCells.size || [...cells].some((c) => !savedCells.has(c));
  const dayOff = disabledWeekdays.includes(wd);

  return (
    <Card>
      <div className="flex items-center justify-between mb-2">
        <h3 className="text-sm font-bold text-neutral-800">매주 반복 시간차단</h3>
        <span className="text-[11px] text-neutral-400">요일별 · 매주 자동 적용</span>
      </div>
      <p className="text-xs text-neutral-500 mb-3">
        “매주 화요일 오후는 항상 막기” 같은 반복 차단. 날짜마다 따로 넣을 필요 없습니다.
      </p>

      {/* 요일 탭 */}
      <div className="flex gap-2 flex-wrap mb-3">
        {DAY_NAMES.map((d, i) => {
          const has = (weeklyMap.get(i) || []).length > 0;
          return (
            <button
              key={i} type="button" onClick={() => setWd(i)}
              className={`relative w-10 h-10 rounded-lg text-sm font-medium transition ${
                wd === i ? 'bg-neutral-900 text-white border border-neutral-900'
                         : 'bg-white text-neutral-600 border border-neutral-200 hover:border-neutral-400'
              }`}
            >
              {d}
              {has && <span className="absolute -top-1 -right-1 w-2 h-2 rounded-full bg-orange-500" />}
            </button>
          );
        })}
      </div>

      {dayOff ? (
        <p className="text-xs text-neutral-400 py-3">
          {DAY_NAMES[wd]}요일은 <b className="text-red-500">정기 휴무</b>라 종일 차단됩니다. (시간 차단 불필요)
        </p>
      ) : (
        <>
          <TimeGrid
            startHour={startHour} endHour={endHour}
            blocked={cells} onChange={setCells}
            disabled={replace.isPending}
          />
          <div className="flex items-center gap-2 mt-3">
            <Button
              size="sm"
              disabled={!dirty || replace.isPending}
              onClick={() => replace.mutate({ weekday: wd, ranges: cellsToRanges(cells) })}
            >
              {replace.isPending ? '저장 중...' : `매주 ${DAY_NAMES[wd]}요일 저장`}
            </Button>
            {dirty && <button onClick={() => setCells(rangesToCells(saved))} className="text-xs text-neutral-400 hover:text-neutral-600">되돌리기</button>}
            {saved.length > 0 && !dirty && (
              <span className="text-[11px] text-orange-600">현재: {summarizeRanges(saved)}</span>
            )}
          </div>
        </>
      )}
    </Card>
  );
}

/**
 * 30분 격자 — 칸을 눌러(드래그) 막기/열기. 구글 캘린더 근무시간·캘린들리 스타일.
 * blocked 에 있는 칸 = 막힘(주황). 나머지 = 열림.
 */
function TimeGrid({ startHour, endHour, blocked, onChange, busyCells, disabled }: {
  startHour: number; endHour: number;
  blocked: Set<number>;
  onChange: (next: Set<number>) => void;
  /** 예약이 이미 있는 칸 (표시만 — 막기 실수 방지) */
  busyCells?: Set<number>;
  disabled?: boolean;
}) {
  const cells = gridCells(startHour, endHour);
  const am = cells.filter((m) => m < 12 * 60);
  const pm = cells.filter((m) => m >= 12 * 60);
  // 드래그 칠하기
  const [drag, setDrag] = useState<null | boolean>(null); // true=막는중, false=여는중

  const apply = (m: number, mode: boolean) => {
    const next = new Set(blocked);
    if (mode) next.add(m); else next.delete(m);
    onChange(next);
  };
  const startDrag = (m: number) => { if (disabled) return; const mode = !blocked.has(m); setDrag(mode); apply(m, mode); };
  const overDrag = (m: number) => { if (disabled || drag === null) return; apply(m, drag); };

  const row = (label: string, list: number[]) => list.length === 0 ? null : (
    <div className="mb-2.5">
      <p className="text-[11px] font-semibold text-neutral-400 mb-1">{label}</p>
      <div className="grid grid-cols-6 sm:grid-cols-8 gap-1">
        {list.map((m) => {
          const on = blocked.has(m);
          const busy = busyCells?.has(m);
          return (
            <button
              key={m} type="button" disabled={disabled}
              onMouseDown={() => startDrag(m)}
              onMouseEnter={() => overDrag(m)}
              onTouchStart={() => startDrag(m)}
              title={busy ? '이 시간에 예약이 있습니다' : undefined}
              className={`h-9 rounded-lg text-[11px] font-semibold border transition select-none disabled:opacity-50 ${
                on ? 'bg-orange-500 border-orange-500 text-white'
                   : busy ? 'bg-emerald-50 border-emerald-300 text-emerald-700'
                   : 'bg-white border-neutral-200 text-neutral-500 hover:border-neutral-400'
              }`}
            >{toHHMM(m)}</button>
          );
        })}
      </div>
    </div>
  );

  return (
    <div onMouseUp={() => setDrag(null)} onMouseLeave={() => setDrag(null)} onTouchEnd={() => setDrag(null)}>
      {row('오전', am)}
      {row('오후', pm)}
      <p className="text-[11px] text-neutral-400 mt-1">
        칸을 누르거나 드래그해서 <b className="text-orange-600">막기</b>/열기 · <span className="text-emerald-600">초록</span>=예약 있음
      </p>
    </div>
  );
}

function pad2(n: number) {
  return String(n).padStart(2, '0');
}

function ymd(year: number, month: number, day: number): string {
  return `${year}-${pad2(month + 1)}-${pad2(day)}`;
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

export default function CalendarManagePage() {
  const today = new Date();
  const todayStr = ymd(today.getFullYear(), today.getMonth(), today.getDate());

  // 사장님 4개월: 현재월 ~ +3개월
  const months = useMemo(() => {
    const arr: { year: number; month: number }[] = [];
    for (let i = 0; i < 4; i++) {
      const d = new Date(today.getFullYear(), today.getMonth() + i, 1);
      arr.push({ year: d.getFullYear(), month: d.getMonth() });
    }
    return arr;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 4개월 전체 범위
  const fromDate = ymd(months[0].year, months[0].month, 1);
  const lastMonth = months[months.length - 1];
  const lastDay = new Date(lastMonth.year, lastMonth.month + 1, 0).getDate();
  const toDate = ymd(lastMonth.year, lastMonth.month, lastDay);

  const { data, isLoading } = useBlackouts(fromDate, toDate);
  const { data: blockedData } = useBlockedSlots(fromDate, toDate);
  const { data: cs } = useConsultationSettings();
  const updateWeekdays = useUpdateDisabledWeekdays();
  const updateBizHours = useUpdateBusinessHours();

  const disabledWeekdays = cs?.disabled_weekdays ?? [0];

  // 영업 기본시간 — 로컬 입력 상태 (cs 로드 시 동기화)
  const [bizStart, setBizStart] = useState(10);
  const [bizEnd, setBizEnd] = useState(20);
  useEffect(() => {
    if (cs) { setBizStart(cs.start_hour ?? 10); setBizEnd(cs.end_hour ?? 20); }
  }, [cs]);
  const bizDirty = cs ? (bizStart !== cs.start_hour || bizEnd !== cs.end_hour) : false;

  const toggleWeekday = (i: number) => {
    const next = disabledWeekdays.includes(i)
      ? disabledWeekdays.filter((w) => w !== i)
      : [...disabledWeekdays, i];
    updateWeekdays.mutate(next);
  };

  // 데이터 → 빠른 조회용 Set/Map
  const blackoutSet = useMemo(() => {
    const s = new Set<string>();
    for (const b of data?.blackouts || []) s.add(b.date);
    return s;
  }, [data]);

  const blackoutReasonMap = useMemo(() => {
    const m = new Map<string, string | null>();
    for (const b of data?.blackouts || []) m.set(b.date, b.reason);
    return m;
  }, [data]);

  const consultMap = useMemo(() => {
    const m = new Map<string, BlackoutConsultation[]>();
    for (const c of data?.consultations || []) {
      if (!c.visit_date) continue;
      if (!m.has(c.visit_date)) m.set(c.visit_date, []);
      m.get(c.visit_date)!.push(c);
    }
    return m;
  }, [data]);

  // 096: 날짜별 시간대 차단 Map (date 행만)
  const blockedSlotMap = useMemo(() => {
    const m = new Map<string, BlockedSlot[]>();
    for (const b of blockedData?.blockedSlots || []) {
      if (!b.date) continue;
      if (!m.has(b.date)) m.set(b.date, []);
      m.get(b.date)!.push(b);
    }
    return m;
  }, [blockedData]);

  // 118: 매주 반복 시간차단 Map (weekday 행만)
  const weeklyBlockedMap = useMemo(() => {
    const m = new Map<number, BlockedSlot[]>();
    for (const b of blockedData?.blockedSlots || []) {
      if (b.weekday == null) continue;
      if (!m.has(b.weekday)) m.set(b.weekday, []);
      m.get(b.weekday)!.push(b);
    }
    for (const list of m.values()) list.sort((a, b) => a.start_time.localeCompare(b.start_time));
    return m;
  }, [blockedData]);

  const [selectedDate, setSelectedDate] = useState<string | null>(null);

  return (
    <>
      <Topbar title="달력 관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 안내 카드 */}
        <Card>
          <div className="flex items-start gap-2">
            <Info size={16} className="shrink-0 mt-0.5 text-blue-500" />
            <div className="text-xs text-neutral-600 space-y-1">
              <p>
                <span className="font-bold text-neutral-800">막힌 날짜·시간은 모든 고객 셀프예약(매장/출장/톡상담 + 복원수리 직접방문)에서 비활성화</span>됩니다.
                사장님이 직접 등록(일정수동등록/시간제안)할 때는 막힌 날짜·시간에도 자유롭게 등록 가능합니다.
              </p>
              <p>
                현재월부터 4개월(<span className="font-medium">{months[0].year}년 {months[0].month + 1}월 ~ {lastMonth.year}년 {lastMonth.month + 1}월</span>)을 관리할 수 있습니다.
                고객은 이번달부터 3개월(<span className="font-medium">{months[0].year}년 {months[0].month + 1}월 ~ {months[2].year}년 {months[2].month + 1}월</span>)까지 예약 가능.
              </p>
            </div>
          </div>
        </Card>

        {/* 영업 기본시간 (096: 설정 탭에서 이전) */}
        <Card>
          <div className="flex items-center justify-between mb-2">
            <div className="flex items-center gap-1.5">
              <Clock size={14} className="text-neutral-400" />
              <h3 className="text-sm font-bold text-neutral-800">영업 기본시간</h3>
            </div>
            <span className="text-[11px] text-neutral-400">고객 예약 가능 시간대</span>
          </div>
          <p className="text-xs text-neutral-500 mb-3">모든 고객 셀프예약 폼의 가용 시간 슬롯 범위에 반영됩니다.</p>
          <div className="flex items-center gap-2 flex-wrap">
            <input
              type="number" min={0} max={23} value={bizStart}
              onChange={(e) => setBizStart(Number(e.target.value))}
              className="w-16 h-9 px-2 rounded-lg border border-stone-200 text-sm text-center focus:outline-none focus:border-stone-400"
            />
            <span className="text-sm text-neutral-600">시 ~</span>
            <input
              type="number" min={0} max={23} value={bizEnd}
              onChange={(e) => setBizEnd(Number(e.target.value))}
              className="w-16 h-9 px-2 rounded-lg border border-stone-200 text-sm text-center focus:outline-none focus:border-stone-400"
            />
            <span className="text-sm text-neutral-600">시</span>
            <Button
              size="sm"
              onClick={() => updateBizHours.mutate({ start_hour: bizStart, end_hour: bizEnd })}
              disabled={!bizDirty || bizStart >= bizEnd || updateBizHours.isPending}
              className="ml-2"
            >
              {updateBizHours.isPending ? '저장 중...' : '저장'}
            </Button>
            {bizStart >= bizEnd && <span className="text-[11px] text-red-500">종료 시간이 시작보다 늦어야 합니다</span>}
          </div>
        </Card>

        {/* 정기 휴무 요일 토글 */}
        <Card>
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-sm font-bold text-neutral-800">정기 휴무 요일</h3>
            <span className="text-[11px] text-neutral-400">매주 반복</span>
          </div>
          <p className="text-xs text-neutral-500 mb-3">선택한 요일은 매주 자동으로 고객 예약 차단됩니다.</p>
          <div className="flex gap-2 flex-wrap">
            {DAY_NAMES.map((d, i) => {
              const isOff = disabledWeekdays.includes(i);
              return (
                <button
                  key={i}
                  type="button"
                  onClick={() => toggleWeekday(i)}
                  disabled={updateWeekdays.isPending}
                  className={`w-10 h-10 rounded-lg text-sm font-medium transition ${
                    isOff
                      ? 'bg-red-500 text-white border border-red-500'
                      : 'bg-white text-neutral-600 border border-neutral-200 hover:border-neutral-400'
                  } ${updateWeekdays.isPending ? 'opacity-50 cursor-wait' : ''}`}
                  title={isOff ? `매주 ${d}요일 휴무 (해제하려면 클릭)` : `${d}요일 정기 휴무로 설정`}
                >
                  {d}
                </button>
              );
            })}
          </div>
        </Card>

        {/* 118: 매주 반복 시간차단 — 요일 골라 시간 격자로 막기 */}
        <WeeklyBlockCard
          startHour={cs?.start_hour ?? 10}
          endHour={cs?.end_hour ?? 20}
          weeklyMap={weeklyBlockedMap}
          disabledWeekdays={disabledWeekdays}
        />

        {/* 4개월 달력 grid (PC 2x2 / 모바일 1열) */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {months.map(({ year, month }, i) => (
            <MonthCalendar
              key={`${year}-${month}`}
              year={year}
              month={month}
              isFirstMonth={i === 0}
              todayStr={todayStr}
              blackoutSet={blackoutSet}
              disabledWeekdays={disabledWeekdays}
              consultMap={consultMap}
              blockedSlotMap={blockedSlotMap}
              weeklyBlockedMap={weeklyBlockedMap}
              onSelect={setSelectedDate}
            />
          ))}
        </div>

        {isLoading && (
          <Card>
            <p className="text-center text-sm text-neutral-400 py-4">로딩 중...</p>
          </Card>
        )}
      </div>

      {selectedDate && (
        <BlackoutDetailModal
          date={selectedDate}
          isBlackout={blackoutSet.has(selectedDate)}
          existingReason={blackoutReasonMap.get(selectedDate) || ''}
          consultations={consultMap.get(selectedDate) || []}
          blockedSlots={blockedSlotMap.get(selectedDate) || []}
          weeklyForThisDay={weeklyBlockedMap.get(new Date(selectedDate + 'T00:00:00').getDay()) || []}
          bizStart={cs?.start_hour ?? 10}
          bizEnd={cs?.end_hour ?? 20}
          isWeekdayOff={disabledWeekdays.includes(new Date(
            Number(selectedDate.slice(0, 4)),
            Number(selectedDate.slice(5, 7)) - 1,
            Number(selectedDate.slice(8, 10)),
          ).getDay())}
          onClose={() => setSelectedDate(null)}
        />
      )}
    </>
  );
}

interface MonthCalendarProps {
  year: number;
  month: number;
  isFirstMonth: boolean;
  todayStr: string;
  blackoutSet: Set<string>;
  disabledWeekdays: number[];
  consultMap: Map<string, BlackoutConsultation[]>;
  blockedSlotMap: Map<string, BlockedSlot[]>;
  weeklyBlockedMap: Map<number, BlockedSlot[]>;
  onSelect: (date: string) => void;
}

function MonthCalendar({
  year, month, isFirstMonth, todayStr, blackoutSet, disabledWeekdays, consultMap, blockedSlotMap, weeklyBlockedMap, onSelect,
}: MonthCalendarProps) {
  const days = getCalendarDays(year, month);

  return (
    <Card>
      <div className="flex items-center gap-2 mb-3">
        <CalendarOff size={14} className="text-neutral-400" />
        <h3 className="text-sm font-bold text-neutral-800">
          {year}년 {month + 1}월
          {!isFirstMonth && month - new Date().getMonth() === 3 && (
            <span className="ml-2 text-[10px] text-amber-600 font-medium">(사장님 전용 — 고객 예약 불가)</span>
          )}
        </h3>
      </div>

      {/* 요일 헤더 */}
      <div className="grid grid-cols-7 gap-1 mb-1">
        {DAY_NAMES.map((d, i) => (
          <div
            key={d}
            className={`text-center text-[11px] font-medium py-1 ${
              i === 0 ? 'text-red-500' : i === 6 ? 'text-blue-500' : 'text-neutral-500'
            }`}
          >
            {d}
          </div>
        ))}
      </div>

      {/* 날짜 grid */}
      <div className="grid grid-cols-7 gap-1">
        {days.map((day, idx) => {
          if (day === null) {
            return <div key={idx} className="aspect-square" />;
          }
          const date = ymd(year, month, day);
          const isPast = date < todayStr;
          const isToday = date === todayStr;
          const isBlackout = blackoutSet.has(date);
          const dow = idx % 7;
          const isWeekdayOff = disabledWeekdays.includes(dow);
          const consults = consultMap.get(date) || [];
          const storeCount = consults.filter((c) => c.consultation_type === 'store_visit').length;
          const fieldCount = consults.filter((c) => c.consultation_type === 'field_request').length;
          // 118: 그 날짜 차단 = 날짜형 + 그 요일의 매주 반복. 실제 시각을 셀에 보여준다
          const dayBlocks = [...(blockedSlotMap.get(date) || []), ...(weeklyBlockedMap.get(dow) || [])];
          const blockedCount = dayBlocks.length;
          const blockedLabel = summarizeRanges(dayBlocks);

          return (
            <button
              key={idx}
              type="button"
              onClick={() => !isPast && onSelect(date)}
              disabled={isPast}
              className={`relative aspect-square rounded-lg border text-xs flex flex-col items-center justify-start py-1 transition ${
                isPast
                  ? 'bg-neutral-50 text-neutral-300 border-neutral-100 cursor-not-allowed'
                  : isBlackout
                  ? 'bg-red-50 border-red-200 hover:border-red-300'
                  : isWeekdayOff
                  ? 'bg-neutral-100 border-neutral-200 hover:border-neutral-400'
                  : isToday
                  ? 'bg-blue-50 border-blue-300 hover:border-blue-400'
                  : 'bg-white border-neutral-200 hover:border-neutral-400'
              }`}
            >
              <span
                className={`font-medium ${
                  isBlackout
                    ? 'text-red-600 line-through'
                    : isWeekdayOff
                    ? 'text-neutral-400 line-through'
                    : dow === 0
                    ? 'text-red-500'
                    : dow === 6
                    ? 'text-blue-500'
                    : 'text-neutral-700'
                }`}
              >
                {day}
              </span>
              {/* 휴무 라벨 (특정날짜 우선, 정기 그 다음) */}
              {isBlackout ? (
                <span className="text-[9px] text-red-500 font-semibold mt-0.5">휴무</span>
              ) : isWeekdayOff ? (
                <span className="text-[9px] text-neutral-400 font-medium mt-0.5">정기휴무</span>
              ) : null}
              {/* 118: 시간차단은 몇시~몇시인지 항상 보이게 (휴무일이어도 가리지 않음) */}
              {blockedCount > 0 && (
                <span
                  className={`text-[9px] font-semibold mt-0.5 leading-tight px-0.5 ${
                    isBlackout || isWeekdayOff ? 'text-orange-400/80' : 'text-orange-600'
                  }`}
                  title={dayBlocks.map((b) => `${b.start_time}~${b.end_time}`).join(', ')}
                >
                  {blockedLabel}
                </span>
              )}
              {/* 예약 dot + 시간차단 dot */}
              {!isBlackout && !isWeekdayOff && (storeCount > 0 || fieldCount > 0 || blockedCount > 0) && (
                <div className="flex gap-0.5 mt-auto mb-1">
                  {storeCount > 0 && (
                    <span
                      className="w-1.5 h-1.5 rounded-full bg-emerald-500"
                      title={`매장방문 ${storeCount}건`}
                    />
                  )}
                  {fieldCount > 0 && (
                    <span
                      className="w-1.5 h-1.5 rounded-full bg-purple-500"
                      title={`출장요청 ${fieldCount}건`}
                    />
                  )}
                  {blockedCount > 0 && (
                    <span
                      className="w-1.5 h-1.5 rounded-full bg-orange-500"
                      title={`시간대 차단 ${blockedCount}건`}
                    />
                  )}
                </div>
              )}
            </button>
          );
        })}
      </div>

      {/* 범례 */}
      <div className="flex items-center gap-3 mt-3 text-[10px] text-neutral-400 flex-wrap">
        <span className="flex items-center gap-1">
          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500" /> 매장
        </span>
        <span className="flex items-center gap-1">
          <span className="w-1.5 h-1.5 rounded-full bg-purple-500" /> 출장
        </span>
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded bg-red-100 border border-red-200" /> 임시 휴무
        </span>
        <span className="flex items-center gap-1">
          <span className="w-2 h-2 rounded bg-neutral-100 border border-neutral-200" /> 정기 휴무
        </span>
        <span className="flex items-center gap-1">
          <span className="w-1.5 h-1.5 rounded-full bg-orange-500" /> 시간 차단
        </span>
      </div>
    </Card>
  );
}

interface BlackoutDetailModalProps {
  date: string;
  isBlackout: boolean;
  existingReason: string;
  consultations: BlackoutConsultation[];
  blockedSlots: BlockedSlot[];
  /** 118: 이 날짜 요일의 매주 반복 차단 (표시용) */
  weeklyForThisDay: BlockedSlot[];
  bizStart: number;
  bizEnd: number;
  isWeekdayOff: boolean;
  onClose: () => void;
}

function BlackoutDetailModal({
  date, isBlackout, existingReason, consultations, blockedSlots, weeklyForThisDay, bizStart, bizEnd, isWeekdayOff, onClose,
}: BlackoutDetailModalProps) {
  const [reason, setReason] = useState(existingReason);
  const create = useCreateBlackout();
  const remove = useDeleteBlackout();

  // 118: 시간 격자 (096 날짜 차단을 칸으로 편집)
  const replaceSlots = useReplaceBlockedSlots();
  const savedCells = useMemo(() => rangesToCells(blockedSlots), [blockedSlots]);
  const [gridCellsState, setGridCellsState] = useState<Set<number>>(savedCells);
  useEffect(() => { setGridCellsState(rangesToCells(blockedSlots)); }, [blockedSlots]);
  const gridDirty = gridCellsState.size !== savedCells.size || [...gridCellsState].some((c) => !savedCells.has(c));

  // 이 날짜에 이미 잡힌 예약 시각 → 격자에 초록 표시 (막기 실수 방지)
  const busyCells = useMemo(() => {
    const s = new Set<number>();
    for (const c of consultations) {
      if (!c.visit_time) continue;
      const m = toMin(String(c.visit_time).slice(0, 5));
      s.add(Math.floor(m / SLOT) * SLOT);
    }
    return s;
  }, [consultations]);

  // 날짜 표시 — 한국어 포맷
  const dateLabel = (() => {
    const [y, m, d] = date.split('-').map(Number);
    const dayOfWeek = new Date(y, m - 1, d).getDay();
    return `${y}년 ${m}월 ${d}일 (${DAY_NAMES[dayOfWeek]})`;
  })();

  const handleBlackout = () => {
    create.mutate(
      { date, reason: reason.trim() || undefined },
      { onSuccess: () => onClose() },
    );
  };

  const handleUnblackout = () => {
    remove.mutate(date, { onSuccess: () => onClose() });
  };

  return (
    <Modal open={true} onClose={onClose} title={dateLabel} className="max-w-2xl">
      <div className="space-y-4">
        {/* 정기 휴무 요일 안내 (해당하면) */}
        {isWeekdayOff && !isBlackout && (
          <div className="flex items-start gap-2 p-2 rounded-lg bg-neutral-100 text-xs text-neutral-600">
            <Info size={14} className="shrink-0 mt-0.5" />
            <span>
              이 날짜는 <span className="font-bold">정기 휴무 요일</span>입니다 (매주 반복).
              위 “정기 휴무 요일” 토글에서 해제하면 매주 영향. 이 날짜만 임시로 막으려면 아래 “이 날짜 막기” 버튼 사용 (사유: 임시 추가 휴무).
            </span>
          </div>
        )}

        {/* 기존 예약 list */}
        {consultations.length > 0 && (
          <div className="space-y-2">
            <p className="text-xs font-bold text-neutral-500">이 날짜 예약 ({consultations.length}건)</p>
            <div className="space-y-1">
              {consultations.map((c) => (
                <div
                  key={c.id}
                  className="flex items-center gap-2 p-2 rounded-lg bg-neutral-50 text-xs"
                >
                  {c.consultation_type === 'store_visit' ? (
                    <Store size={12} className="text-emerald-500 shrink-0" />
                  ) : (
                    <Truck size={12} className="text-purple-500 shrink-0" />
                  )}
                  <span className="font-medium text-neutral-800">{c.name}</span>
                  <span className="text-neutral-500">{formatPhone(c.phone)}</span>
                  {c.visit_time && <span className="text-neutral-400">{c.visit_time}</span>}
                  <span className="ml-auto text-[10px] px-1.5 py-0.5 rounded bg-white border border-neutral-200 text-neutral-600">
                    {c.status}
                  </span>
                </div>
              ))}
            </div>
            {!isBlackout && (
              <div className="flex items-start gap-2 p-2 rounded-lg bg-amber-50 text-xs text-amber-700">
                <AlertTriangle size={14} className="shrink-0 mt-0.5" />
                <span>이날 예약이 있습니다. 막아도 기존 예약은 유지되며, 신규 고객 예약만 차단됩니다.</span>
              </div>
            )}
          </div>
        )}

        {/* 휴무 사유 입력 (휴무 등록 모드일 때만) */}
        {!isBlackout && (
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">휴무 사유 (선택)</label>
            <input
              type="text"
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              placeholder="예: 여름휴가, 공휴일, 외부 일정"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
              autoFocus
            />
          </div>
        )}

        {/* 기존 휴무 사유 표시 */}
        {isBlackout && existingReason && (
          <div className="p-2 rounded-lg bg-red-50 text-xs">
            <span className="font-bold text-red-700">사유:</span>{' '}
            <span className="text-red-600">{existingReason}</span>
          </div>
        )}

        {/* 118: 시간 격자 — 칸을 칠해서 막기/열기. 전일 휴무여도 조회·수정 가능(기존엔 숨겨져 손도 못 댔음) */}
        <div className="border-t border-neutral-100 pt-3 space-y-2">
          <div className="flex items-center gap-1.5">
            <Ban size={13} className="text-orange-500" />
            <span className="text-xs font-bold text-neutral-700">이 날짜 시간 열기/막기</span>
            {gridDirty && <span className="ml-auto text-[11px] text-orange-600 font-semibold">저장 안 됨</span>}
          </div>
          <p className="text-[11px] text-neutral-400">
            막은 시간은 모든 셀프예약(매장/출장/톡 + 복원수리 직접방문)에서 비활성. 사장님 직접 등록은 항상 가능.
          </p>
          {isBlackout && (
            <p className="text-[11px] text-red-500">이 날짜는 <b>종일 휴무</b>라 어차피 전부 막힙니다. (휴무 해제 시 아래 설정이 적용됩니다)</p>
          )}
          {weeklyForThisDay.length > 0 && (
            <p className="text-[11px] text-neutral-500">
              매주 {DAY_NAMES[new Date(date + 'T00:00:00').getDay()]}요일 반복 차단({summarizeRanges(weeklyForThisDay)})이 이 날짜에도 적용됩니다.
              <span className="text-neutral-400"> 반복은 위 “매주 반복 시간차단”에서 수정</span>
            </p>
          )}

          <TimeGrid
            startHour={bizStart} endHour={bizEnd}
            blocked={gridCellsState} onChange={setGridCellsState}
            busyCells={busyCells}
            disabled={replaceSlots.isPending}
          />

          <div className="flex items-center gap-2">
            <Button
              size="sm"
              disabled={!gridDirty || replaceSlots.isPending}
              onClick={() => replaceSlots.mutate({ date, ranges: cellsToRanges(gridCellsState) })}
            >
              {replaceSlots.isPending ? '저장 중...' : '시간 설정 저장'}
            </Button>
            {gridDirty && (
              <button onClick={() => setGridCellsState(rangesToCells(blockedSlots))} className="text-xs text-neutral-400 hover:text-neutral-600">되돌리기</button>
            )}
            {gridCellsState.size > 0 && !gridDirty && (
              <span className="text-[11px] text-orange-600">차단: {summarizeRanges(cellsToRanges(gridCellsState))}</span>
            )}
            {gridCellsState.size === 0 && !gridDirty && (
              <span className="text-[11px] text-neutral-400">이 날짜는 종일 예약 가능</span>
            )}
          </div>
        </div>

        {/* 액션 */}
        <div className="flex justify-end gap-2 pt-2">
          <Button variant="ghost" size="sm" onClick={onClose}>닫기</Button>
          {isBlackout ? (
            <Button
              size="sm"
              variant="secondary"
              onClick={handleUnblackout}
              disabled={remove.isPending}
            >
              {remove.isPending ? '해제 중...' : '휴무 해제'}
            </Button>
          ) : (
            <Button
              size="sm"
              onClick={handleBlackout}
              disabled={create.isPending}
            >
              {create.isPending ? '등록 중...' : '이 날짜 막기'}
            </Button>
          )}
        </div>
      </div>
    </Modal>
  );
}
