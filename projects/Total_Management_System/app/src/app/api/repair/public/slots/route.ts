import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { getVisitDay, buildSlots } from '@/lib/repair/visit-schedule';
import { visitDurationMin } from '@/lib/repair/intake';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
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

    // 2026-10-06: 영업설정·휴무·점유 구간 계산은 공용(lib/repair/visit-schedule) — TMS 방문 예약 등록과 같은 규칙
    // 소요시간(점유) = 10분 + 자루당 5분 (submit 의 visit_duration_min 과 동일 공식)
    const blockMin: number = visitDurationMin(qty);
    const day = await getVisitDay(dbAny, date);
    if (day.closedReason) {
      return NextResponse.json(
        { ok: true, slots: [], blockMin, qty, closedDay: true, reason: day.closedReason },
        { headers: CORS_HEADERS }
      );
    }
    // 🔒 고객 페이지엔 시간·가능 여부만 — 겹치는 일정 라벨(다른 고객 이름 포함)은 내보내지 않는다
    const slots = buildSlots(day, date, blockMin).map((x) => ({ time: x.time, available: x.available }));

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
