import { NextRequest, NextResponse } from 'next/server';
import { upsertRepair, batchUpsertRepairs, type RepairGasPushPayload } from '@/lib/repair/sync';

/** POST /api/repair/sync — GAS → TMS 동기화 (접수 데이터 push) */
export async function POST(req: NextRequest) {
  try {
    // 인증: CRON_SECRET 또는 GAS 토큰
    const authHeader = req.headers.get('authorization');
    const cronSecret = process.env.CRON_SECRET || 'mamoru-tms-cron-2026';

    if (authHeader !== `Bearer ${cronSecret}`) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const body = await req.json();

    // 단건 / 배치 분기
    if (Array.isArray(body)) {
      const result = await batchUpsertRepairs(body as RepairGasPushPayload[]);
      return NextResponse.json(result);
    }

    // 단건
    const result = await upsertRepair(body as RepairGasPushPayload);
    if (!result.success) {
      return NextResponse.json({ error: result.error }, { status: 400 });
    }

    return NextResponse.json(result, { status: 201 });
  } catch (err) {
    console.error('[repair/sync] 동기화 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
