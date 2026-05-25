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

/** 분 단위 변환 */
function toMinutes(hhmm: string): number {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}
function fromMinutes(min: number): string {
  const h = Math.floor(min / 60);
  const m = min % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

/** 두 시간 범위가 겹치는지 */
function overlaps(a0: number, a1: number, b0: number, b1: number): boolean {
  return a0 < b1 && a1 > b0;
}

/**
 * GET /api/repair/public/slots?date=YYYY-MM-DD&qty=N
 *
 * 복원수리 직접방문(당일수리) 예약 가능 시간 슬롯 조회 (비인증, CORS).
 *
 * 동작:
 *   1. consultation_settings 의 영업시간 + repair_* 정책 로드
 *   2. 30분 간격 후보 슬롯 생성
 *   3. 같은 날 충돌 데이터 3종 병렬 조회:
 *      - 컨설팅 매장방문 (consultations, consultation_type != 'field_request')
 *      - 컨설팅 출장 (consultations, consultation_type='field_request', buffer 적용)
 *      - 복원수리 직접방문 (repairs, proceed_type='직접방문')
 *   4. qty 기반 차단 시간 결정 (1~5자루=30분 / 6자루+=60분, 운영 정책 컬럼화)
 *   5. 슬롯별 available 판정
 *
 * 응답:
 *   { ok: true, slots: [{ time, available }, ...], blockMin, qty }
 */
export async function GET(req: NextRequest) {
  try {
    const sp = req.nextUrl.searchParams;
    const date = sp.get('date');
    const qty = parseInt(sp.get('qty') || '1', 10);

    // 입력 검증
    if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
      return NextResponse.json(
        { ok: false, error: 'date 파라미터가 필요합니다 (YYYY-MM-DD)' },
        { status: 400, headers: CORS_HEADERS }
      );
    }
    if (!qty || qty < 1) {
      return NextResponse.json(
        { ok: false, error: 'qty 파라미터가 필요합니다 (1 이상)' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 1. consultation_settings 로드 (영업시간 + 운영 정책)
    const { data: settings } = await dbAny
      .from('consultation_settings')
      .select('*')
      .eq('id', 'default')
      .single();

    const startHour: number = settings?.start_hour ?? 10;
    const endHour: number = settings?.end_hour ?? 20;
    const fieldBufferBefore: number = settings?.field_buffer_before ?? 60;
    const fieldBufferAfter: number = settings?.field_buffer_after ?? 60;
    const consultDurMin: number = settings?.duration_min ?? 60;       // 매장방문/출장 차단 길이
    const disabledWeekdays: number[] = settings?.disabled_weekdays ?? [0];

    // 092 신규 정책 컬럼
    const slotStep: number = settings?.repair_slot_step_min ?? 30;
    const thresholdQty: number = settings?.repair_threshold_qty ?? 6;
    const blockUnder: number = settings?.repair_block_under_min ?? 30;
    const blockOver: number = settings?.repair_block_over_min ?? 60;

    // 차단 시간 결정 (qty 기준)
    const blockMin: number = qty >= thresholdQty ? blockOver : blockUnder;

    // 2. 휴무일/요일 체크
    const dayOfWeek = new Date(date + 'T00:00:00').getDay();
    if (disabledWeekdays.includes(dayOfWeek)) {
      return NextResponse.json(
        { ok: true, slots: [], blockMin, qty, closedDay: true, reason: 'disabled_weekday' },
        { headers: CORS_HEADERS }
      );
    }

    const { data: closedDate } = await dbAny
      .from('closed_dates')
      .select('date')
      .eq('date', date)
      .maybeSingle();

    if (closedDate) {
      return NextResponse.json(
        { ok: true, slots: [], blockMin, qty, closedDay: true, reason: 'closed_date' },
        { headers: CORS_HEADERS }
      );
    }

    // 3. 같은 날 충돌 데이터 3종 병렬 조회
    const [consultsRes, repairsRes] = await Promise.all([
      // 컨설팅 매장방문 + 출장 (visit_date 기준, 확정/배정 상태만)
      dbAny.from('consultations')
        .select('consultation_type, visit_time, status')
        .eq('visit_date', date)
        .in('status', ['confirmed', 'assigned']),
      // 복원수리 직접방문 (visit_date 기준, 취소 제외)
      dbAny.from('repairs')
        .select('visit_time, visit_duration_min, status')
        .eq('visit_date', date)
        .eq('proceed_type', '직접방문')
        .neq('status', 'cancelled'),
    ]);

    // 4. 차단 범위 수집 [start_min, end_min)
    const busy: Array<[number, number]> = [];

    for (const c of (consultsRes.data || [])) {
      if (!c.visit_time) continue;
      const t = toMinutes(c.visit_time);
      if (c.consultation_type === 'field_request') {
        // 출장: 전후 버퍼 적용
        busy.push([t - fieldBufferBefore, t + consultDurMin + fieldBufferAfter]);
      } else {
        // 매장방문: 상담 시간만
        busy.push([t, t + consultDurMin]);
      }
    }

    for (const r of (repairsRes.data || [])) {
      if (!r.visit_time) continue;
      const t = toMinutes(r.visit_time);
      const d = (r.visit_duration_min as number) || 30;
      busy.push([t, t + d]);
    }

    // 5. 후보 슬롯 생성 + available 판정
    const startMin = startHour * 60;
    const endMin = endHour * 60;
    const today = new Date().toISOString().slice(0, 10);
    const now = new Date();
    const nowMin = now.getHours() * 60 + now.getMinutes();
    const isToday = date === today;

    const slots: Array<{ time: string; available: boolean }> = [];
    for (let m = startMin; m + blockMin <= endMin; m += slotStep) {
      // 오늘이면 현재 시각 이전 슬롯 비활성
      if (isToday && m <= nowMin) {
        slots.push({ time: fromMinutes(m), available: false });
        continue;
      }
      // 차단 범위와 겹치는지 검사
      const slotEnd = m + blockMin;
      const conflict = busy.some(([b0, b1]) => overlaps(m, slotEnd, b0, b1));
      slots.push({ time: fromMinutes(m), available: !conflict });
    }

    return NextResponse.json(
      { ok: true, slots, blockMin, qty },
      { headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[repair/public/slots] 실패:', err);
    return NextResponse.json(
      { ok: false, error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
