import { NextRequest, NextResponse } from 'next/server';
import { sweepShippingTodos, kstHour } from '@/lib/google/shipping-todo-sync';

const CRON_SECRET = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';
const CREATE_HOUR_KST = 8;   // 아침 8시에만 '할 일' 생성 — 정리(삭제)는 매시간

/**
 * GET /api/cron/shipping-todo — 송장 미생성 건을 구글 캘린더 종일 '할 일'로 (2026-09-24)
 *
 * Vercel Cron 매시간. 한 크론에 생성·정리를 같이 둔 이유:
 *   · 생성은 하루 1회(아침 8시)면 충분하다 — 등록 다음 날부터 뜨는 규칙이라 시각이 중요치 않음
 *   · 정리는 자주 돌아야 한다 — 송장을 만들자마자 캘린더에서 사라져야 믿고 쓴다 (최대 1시간 지연)
 *   판매/납품 API 각각에 삭제 훅을 심지 않고 여기서 일괄 처리 → 경로 누락으로 유령 일정이 남는 일이 없다
 *
 * 수동 확인: ?dry=1 (캘린더를 건드리지 않고 대상만 반환) · ?force=1 (시각 무시하고 생성) · ?asOf=YYYY-MM-DD (기준일 대체)
 */
export async function GET(req: NextRequest) {
  const authHeader = req.headers.get('authorization');
  if (authHeader !== `Bearer ${CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const sp = req.nextUrl.searchParams;
  const dryRun = sp.get('dry') === '1';
  const create = sp.get('force') === '1' || kstHour() === CREATE_HOUR_KST;

  const asOf = sp.get('asOf') || undefined;

  const result = await sweepShippingTodos({ create, dryRun, asOf });
  return NextResponse.json({ ok: !result.error, ...result, createWindow: create });
}
