import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { runOpsWatch } from '@/lib/ops/watch';

/**
 * GET /api/cron/ops-watch — 10분마다 (2026-10-06)
 * Make 시나리오 꺼짐(매번) · Make 오류(1시간) · 아침 점검(08시 KST). 상세 규칙은 lib/ops/watch.ts
 *
 * 수동 확인: ?dry=1 (발송·상태저장 없이 결과만) · ?force=hourly|daily (시각 무시하고 실행)
 */
export async function GET(req: NextRequest) {
  if (req.headers.get('authorization') !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  try {
    const sp = req.nextUrl.searchParams;
    const f = sp.get('force');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const result = await runOpsWatch(db, {
      dryRun: sp.get('dry') === '1',
      force: f === 'hourly' || f === 'daily' ? f : undefined,
    });
    if (result.alerts.length || result.errors.length) console.log('[cron/ops-watch]', JSON.stringify(result));
    return NextResponse.json(result);
  } catch (err) {
    console.error('[cron/ops-watch] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
