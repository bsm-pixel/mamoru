import { NextRequest, NextResponse } from 'next/server';
import { syncOrders } from '@/lib/imweb/sync';

/** GET /api/cron/sync-orders — 자동 동기화 (Vercel Cron) */
export async function GET(request: NextRequest) {
  // Cron 시크릿 검증
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const result = await syncOrders();
    return NextResponse.json(result);
  } catch (err) {
    console.error('[cron] 동기화 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
