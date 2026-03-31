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
    const fieldBufferBefore = settings?.field_buffer_before ?? 90;
    const fieldBufferAfter = settings?.field_buffer_after ?? 90;
    const disabledWeekdays: number[] = settings?.disabled_weekdays ?? [0];

    // 2. 휴무일 조회
    const { data: closedDatesData } = await dbAny
      .from('closed_dates')
      .select('date')
      .in('date', dates);
    const closedSet = new Set((closedDatesData || []).map((d: { date: string }) => d.date));

    // 3. 해당 날짜 범위의 확정/제안 상담 조회 (슬롯 차단용)
    const minDate = dates.reduce((a, b) => a < b ? a : b);
    const maxDate = dates.reduce((a, b) => a > b ? a : b);

    const { data: bookings } = await dbAny
      .from('consultations')
      .select('consultation_type, visit_date, visit_time, status, suggestions')
      .gte('visit_date', minDate)
      .lte('visit_date', maxDate)
      .in('status', ['confirmed', 'suggested', 'assigned', 'pending_admin']);

    // 4. 날짜별 차단 슬롯 계산
    const blockedMap = new Map<string, Set<number>>();

    for (const b of (bookings || [])) {
      if (!b.visit_date || !b.visit_time) continue;

      const dateStr = b.visit_date;
      if (!blockedMap.has(dateStr)) blockedMap.set(dateStr, new Set());
      const blocked = blockedMap.get(dateStr)!;

      const baseMin = toMinutes(b.visit_time);

      if (b.consultation_type === 'field_request' && ['suggested', 'assigned'].includes(b.status)) {
        // 출장: 전후 버퍼 적용
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

      // suggested 상태의 제안 시간도 차단
      if (b.status === 'suggested' && b.suggestions) {
        const sug = b.suggestions as Array<{ date?: string; time?: string }>;
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

      result[dateStr] = allSlots.filter(slot => {
        const slotMin = toMinutes(slot);
        // 오늘이면 현재 시각 이전 슬롯 제외
        if (dateStr === today && slotMin <= nowMin) return false;
        // 차단된 슬롯인지 확인 (슬롯의 시작~종료 범위가 차단 범위와 겹치는지)
        for (let m = slotMin; m < slotMin + durMin; m += stepMin) {
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
