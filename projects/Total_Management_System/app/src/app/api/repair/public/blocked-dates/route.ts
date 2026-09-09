import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

function fmt(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * GET /api/repair/public/blocked-dates — 관리자가 지정한 '수거 불가일' (비인증)
 * 설정 키 repair.pickup_blocked_dates = [{start, end, reason}] 를 개별 날짜로 펼쳐 반환.
 * 고객 접수폼/수거일 변경 달력에서 이 날짜를 선택 불가 처리.
 */
export async function GET() {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const { data, error } = await dbAny
      .from('settings')
      .select('value')
      .eq('key', 'repair.pickup_blocked_dates')
      .limit(1);
    if (error) throw new Error(error.message || error.hint || JSON.stringify(error));

    const row = Array.isArray(data) && data.length ? data[0] : null;
    let ranges: { start?: string; end?: string; reason?: string }[] = [];
    if (row?.value) {
      try {
        ranges = typeof row.value === 'string' ? JSON.parse(row.value) : row.value;
      } catch {
        ranges = [];
      }
    }
    if (!Array.isArray(ranges)) ranges = [];

    const dates = new Set<string>();
    const reasons: Record<string, string> = {};
    for (const r of ranges) {
      if (!r || !r.start) continue;
      const start = new Date(r.start + 'T00:00:00');
      const end = new Date((r.end || r.start) + 'T00:00:00');
      if (isNaN(start.getTime()) || isNaN(end.getTime())) continue;
      // 안전장치: 최대 400일까지만 펼침
      let guard = 0;
      for (const d = new Date(start); d <= end && guard < 400; d.setDate(d.getDate() + 1), guard++) {
        const ds = fmt(d);
        dates.add(ds);
        if (r.reason) reasons[ds] = r.reason;
      }
    }

    return NextResponse.json(
      { ok: true, dates: [...dates], reasons, ranges },
      { headers: CORS_HEADERS },
    );
  } catch (err) {
    console.error('[repair/public/blocked-dates] 조회 실패:', err);
    const msg = err instanceof Error ? err.message : (typeof err === 'string' ? err : JSON.stringify(err));
    return NextResponse.json(
      { ok: false, error: msg, dates: [], reasons: {}, ranges: [] },
      { status: 500, headers: CORS_HEADERS },
    );
  }
}
