import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { runExpenseMailScan } from '@/lib/expenses/mail-scan';

/**
 * GET /api/cron/expense-mail — 매일 08시(KST) (2026-10-29)
 * 영수증·세금계산서 메일 → 경비 「확인 대기」. 규칙은 lib/expenses/mail-parse.ts
 * 수동 확인: ?days=120 (처음 한 번 과거분 모으기)
 */
export async function GET(req: NextRequest) {
  if (req.headers.get('authorization') !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  try {
    const days = Math.min(365, Math.max(1, parseInt(req.nextUrl.searchParams.get('days') || '3', 10) || 3));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const result = await runExpenseMailScan(db, days);
    console.log('[cron/expense-mail]', JSON.stringify(result));
    return NextResponse.json(result);
  } catch (err) {
    console.error('[cron/expense-mail] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
