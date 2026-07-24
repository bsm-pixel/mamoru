import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** 시간 슬롯 생성: startHour~endHour, stepMin 간격 */
function generateTimeSlots(startHour: number, endHour: number, stepMin: number, durMin: number): string[] {
  const slots: string[] = [];
  const startMin = startHour * 60;
  const endMin = endHour * 60;
  // 마지막 슬롯 = endHour - duration (예: 20시 영업종료, 60분 상담 → 19:00이 마지막)
  const lastSlotMin = endMin - durMin;
  for (let m = startMin; m <= lastSlotMin; m += stepMin) {
    const h = String(Math.floor(m / 60)).padStart(2, '0');
    const min = String(m % 60).padStart(2, '0');
    slots.push(`${h}:${min}`);
  }
  return slots;
}

/** 분 단위로 변환 */
function toMinutes(time: string): number {
  const [h, m] = time.split(':').map(Number);
  return h * 60 + m;
}

/** 두 시간 범위가 겹치는지 */
function overlaps(a0: number, a1: number, b0: number, b1: number): boolean {
  return a0 < b1 && a1 > b0;
}

/**
 * GET /api/consultation/public/slots?dates=2026-04-01,2026-04-02&type=매장 방문
 * 예약 가능 시간 슬롯 조회 (비인증, CORS)
 *
 * GAS getSlotsRange() + getSlotsMonth_SheetBased_() 로직 이전
 */
export async function GET(req: NextRequest) {
  try {
    const datesParam = req.nextUrl.searchParams.get('dates');
    if (!datesParam) {
      return NextResponse.json(
        { ok: false, error: 'dates 파라미터가 필요합니다 (YYYY-MM-DD 콤마 구분)' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const dates = datesParam.split(',').map(d => d.trim()).filter(Boolean);
    if (dates.length === 0 || dates.length > 62) {
      return NextResponse.json(
        { ok: false, error: '날짜는 1~62개까지 조회 가능합니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 1. 설정 조회
    const { data: settings } = await dbAny
      .from('consultation_settings')
      .select('*')
      .eq('id', 'default')
      .single();

    const startHour = settings?.start_hour ?? 10;
    const endHour = settings?.end_hour ?? 20;
    const durMin = settings?.duration_min ?? 60;
    const stepMin = settings?.step_min ?? 10;
    const fieldBufferBefore = settings?.field_buffer_before ?? 60; // 이동 1h (상담종료 후 출발)
    const fieldBufferAfter = settings?.field_buffer_after ?? 60; // 복귀 1h (상담시간 durMin은 별도 합산)
    const disabledWeekdays: number[] = settings?.disabled_weekdays ?? [0];

    // 2. 휴무일 조회
    const { data: closedDatesData } = await dbAny
      .from('closed_dates')
      .select('date')
      .in('date', dates);
    const closedSet = new Set((closedDatesData || []).map((d: { date: string }) => d.date));

    // 2-2. 시간대 차단 조회 (096) — 사장님 개인 일정 등 30분 단위 부분 차단
    const { data: blockedSlotsData } = await dbAny
      .from('blocked_time_slots')
      .select('date, start_time, end_time')
      .in('date', dates);

    // 2-3. 매주 반복 시간차단 (118) — weekday(0=일~6=토) 행. 해당 요일의 모든 날짜에 적용
    const { data: weeklyBlockedData } = await dbAny
      .from('blocked_time_slots')
      .select('weekday, start_time, end_time')
      .not('weekday', 'is', null);

    // 3. 해당 날짜 범위의 확정/제안 상담 조회 (슬롯 차단용)
    const minDate = dates.reduce((a, b) => a < b ? a : b);
    const maxDate = dates.reduce((a, b) => a > b ? a : b);

    // 확정/배정/대기 건: visit_date 범위로 조회
    const { data: dateBookings } = await dbAny
      .from('consultations')
      .select('consultation_type, visit_date, visit_time, status, suggestions')
      .gte('visit_date', minDate)
      .lte('visit_date', maxDate)
      .in('status', ['confirmed', 'assigned', 'pending_admin']);

    // suggested 건: visit_date가 null일 수 있으므로 별도 조회 (suggestions JSONB 안에 날짜 있음)
    const { data: suggestedBookings } = await dbAny
      .from('consultations')
      .select('consultation_type, visit_date, visit_time, status, suggestions')
      .eq('status', 'suggested')
      .not('suggestions', 'is', null);

    const bookings = [...(dateBookings || []), ...(suggestedBookings || [])];

    // 3-2. 복원수리 직접방문 충돌 데이터 조회 (2026-05-25 — 양방향 차단 완성)
    //   직접방문은 visit_date + visit_time + visit_duration_min 사용 (092 마이그레이션)
    //   매장방문/출장 슬롯 검색 시에도 직접방문 일정을 차단해야 양방향 유기적 흐름
    const { data: directVisits } = await dbAny
      .from('repairs')
      .select('visit_date, visit_time, visit_duration_min')
      .gte('visit_date', minDate)
      .lte('visit_date', maxDate)
      .eq('proceed_type', '직접방문')
      .neq('status', 'cancelled');

    // 4. 날짜별 차단 슬롯 계산
    const blockedMap = new Map<string, Set<number>>();

    // 4-0. 시간대 차단 (096 날짜 + 118 매주반복) — 구간(interval) 으로 보관해 슬롯과 '겹침' 판정.
    //   ⚠️ 예전에는 stepMin 간격으로 Set 에 분(minute)을 넣고 정확히 일치할 때만 막았는데,
    //      차단 시작(예 12:30)과 슬롯 격자(startHour 기점 stepMin)의 위상이 어긋나면
    //      (예: step_min=20/60) 차단이 조용히 무효화됐다. 구간겹침이면 격자와 무관하게 항상 막힌다.
    //      (복원수리 슬롯 API 가 쓰던 방식으로 통일 — 2026-07-24)
    const blockedIntervalMap = new Map<string, Array<[number, number]>>();
    const addBlockedInterval = (dateStr: string, st: string, et: string) => {
      const s = toMinutes(st), e = toMinutes(et);
      if (!(e > s)) return;
      if (!blockedIntervalMap.has(dateStr)) blockedIntervalMap.set(dateStr, []);
      blockedIntervalMap.get(dateStr)!.push([s, e]);
    };
    for (const bs of (blockedSlotsData || [])) {
      if (!bs.date || !bs.start_time || !bs.end_time) continue;
      addBlockedInterval(bs.date as string, bs.start_time as string, bs.end_time as string);
    }
    // 매주 반복(118): 각 조회 날짜의 요일과 일치하는 반복 차단을 그 날짜에 전개
    for (const wb of (weeklyBlockedData || [])) {
      if (wb.weekday == null || !wb.start_time || !wb.end_time) continue;
      for (const d of dates) {
        const [yy, mm, dd] = d.split('-').map(Number);
        if (new Date(yy, mm - 1, dd).getDay() !== Number(wb.weekday)) continue;
        addBlockedInterval(d, wb.start_time as string, wb.end_time as string);
      }
    }

    // 4-A. 복원수리 직접방문 차단 (먼저 처리, 단순 시간 + duration 차단)
    for (const v of (directVisits || [])) {
      if (!v.visit_date || !v.visit_time) continue;
      const dateStr = v.visit_date as string;
      if (!blockedMap.has(dateStr)) blockedMap.set(dateStr, new Set());
      const blocked = blockedMap.get(dateStr)!;
      const baseMin = toMinutes(v.visit_time as string);
      const visitDur = (v.visit_duration_min as number) || 30;
      for (let m = baseMin; m < baseMin + visitDur; m += stepMin) {
        blocked.add(m);
      }
    }

    for (const b of (bookings || [])) {
      // suggested 건은 visit_time이 null일 수 있으므로 suggestions만 처리
      if (b.status === 'suggested' && b.suggestions) {
        // suggestions 구조: { dates: [{ date, time }, ...] } 또는 [{ date, time }, ...]
        const raw = b.suggestions as { dates?: Array<{ date?: string; time?: string }> } | Array<{ date?: string; time?: string }>;
        const sug = Array.isArray(raw) ? raw : (raw.dates || []);
        for (const s of sug) {
          if (s.date && s.time) {
            if (!blockedMap.has(s.date)) blockedMap.set(s.date, new Set());
            const sugBlocked = blockedMap.get(s.date)!;
            const sugMin = toMinutes(s.time);
            if (b.consultation_type === 'field_request') {
              const start = sugMin - fieldBufferBefore;
              const end = sugMin + durMin + fieldBufferAfter;
              for (let m = start; m < end; m += stepMin) {
                sugBlocked.add(m);
              }
            } else {
              for (let m = sugMin; m < sugMin + durMin; m += stepMin) {
                sugBlocked.add(m);
              }
            }
          }
        }
        continue; // suggestions 처리 완료, visit_time 기반 차단은 건너뜀
      }

      // visit_date/visit_time이 없는 건은 차단 불가
      if (!b.visit_date || !b.visit_time) continue;

      const dateStr = b.visit_date;
      if (!blockedMap.has(dateStr)) blockedMap.set(dateStr, new Set());
      const blocked = blockedMap.get(dateStr)!;

      const baseMin = toMinutes(b.visit_time);

      if (b.consultation_type === 'field_request' && ['confirmed', 'assigned'].includes(b.status)) {
        // 출장 확정/배정: 전후 버퍼 적용
        const blockStart = baseMin - fieldBufferBefore;
        const blockEnd = baseMin + durMin + fieldBufferAfter;
        for (let m = blockStart; m < blockEnd; m += stepMin) {
          blocked.add(m);
        }
      } else {
        // 매장방문/기타: 상담 시간만 차단
        for (let m = baseMin; m < baseMin + durMin; m += stepMin) {
          blocked.add(m);
        }
      }
    }

    // 5. 전체 가능 슬롯 생성
    const allSlots = generateTimeSlots(startHour, endHour, stepMin, durMin);
    const today = new Date().toISOString().slice(0, 10);
    const now = new Date();
    const nowMin = now.getHours() * 60 + now.getMinutes();

    // 6. 날짜별 가용 슬롯 계산
    const result: Record<string, string[]> = {};

    for (const dateStr of dates) {
      // 과거일 → 빈 슬롯
      if (dateStr < today) {
        result[dateStr] = [];
        continue;
      }

      // 휴무일 → 빈 슬롯
      if (closedSet.has(dateStr)) {
        result[dateStr] = [];
        continue;
      }

      // 휴무 요일 → 빈 슬롯
      const dayOfWeek = new Date(dateStr + 'T00:00:00').getDay();
      if (disabledWeekdays.includes(dayOfWeek)) {
        result[dateStr] = [];
        continue;
      }

      const blocked = blockedMap.get(dateStr) || new Set();
      const blockedIntervals = blockedIntervalMap.get(dateStr) || [];

      result[dateStr] = allSlots.filter(slot => {
        const slotMin = toMinutes(slot);
        const slotEnd = slotMin + durMin;
        // 오늘이면 현재 시각 이전 슬롯 제외
        if (dateStr === today && slotMin <= nowMin) return false;
        // 시간차단(날짜·매주반복) — 구간겹침. 격자 정렬과 무관하게 항상 정확히 막힌다
        if (blockedIntervals.some(([s, e]) => overlaps(slotMin, slotEnd, s, e))) return false;
        // 그 외 차단(예약·직접방문 등)은 기존 분(minute) Set 방식 유지
        for (let m = slotMin; m < slotEnd; m += stepMin) {
          if (blocked.has(m)) return false;
        }
        return true;
      });
    }

    return NextResponse.json({ ok: true, data: result }, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[consultation/public/slots] 조회 실패:', err);
    return NextResponse.json(
      { ok: false, error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
