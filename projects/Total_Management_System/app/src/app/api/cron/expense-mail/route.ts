import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { runExpenseMailScan } from '@/lib/expenses/mail-scan';
import { fillRecurring, thisMonthKst } from '@/lib/expenses/recurring';
import { archivePendingReceipts } from '@/lib/expenses/receipt-archive';

/**
 * GET /api/cron/expense-mail — 매일 08시(KST) (2026-10-29)
 * 영수증·세금계산서 메일 → 경비 「확인 대기」. 규칙은 lib/expenses/mail-parse.ts
 * 수동 확인: ?days=120 (처음 한 번 과거분 모으기)
 * 169: 고정 경비(매달 반복) 이번 달 분도 여기서 자동 등록 — 이미 있으면 건너뜀 (lib/expenses/recurring.ts)
 * 영수증 원본 보관 — 확인 직후 못 한 것·옛 것 재시도 (lib/expenses/receipt-archive.ts)
 */
export async function GET(req: NextRequest) {
  if (req.headers.get('authorization') !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
  try {
    const days = Math.min(365, Math.max(1, parseInt(req.nextUrl.searchParams.get('days') || '3', 10) || 3));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    // 고정 경비는 메일 확인 실패(구글 연결 끊김 등)와 상관없이 먼저
    const { data: rec } = await db.from('recurring_expenses').select('id, category, amount, memo, proof_type').eq('is_active', true);
    const recurring = await fillRecurring(db, rec || [], [thisMonthKst()]);
    const result = await runExpenseMailScan(db, days);
    const receipts = result.ok ? await archivePendingReceipts(db, 30) : null;
    console.log('[cron/expense-mail]', JSON.stringify({ ...result, recurring, receipts }));
    return NextResponse.json({ ...result, recurring, receipts });
  } catch (err) {
    console.error('[cron/expense-mail] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
