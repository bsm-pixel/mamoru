import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { runFallbackSweep } from '@/lib/firebase/push-fallback';

/**
 * GET /api/cron/push-fallback — 1분마다 (2026-09-29, 마이그 155)
 * 2분 지나도 휴대폰이 받지 못한 관리자 알림 → 사장님 메일 1통(bsm@mamoru.kr → Gmail 앱 알림).
 * 실행할 때마다 심장박동을 남겨서, 이 크론이 멈추면 sendPushToAll 이 기다리지 않고 즉시 메일로 간다.
 */
export async function GET(request: NextRequest) {
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const result = await runFallbackSweep(db);
    if (result.sent || result.failed) console.log('[cron/push-fallback]', result);
    return NextResponse.json(result);
  } catch (err) {
    console.error('[cron/push-fallback] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
