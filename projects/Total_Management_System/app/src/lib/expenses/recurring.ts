/**
 * 고정 경비(매달 반복) → 경비 자동 등록 (2026-10-29, 마이그 169)
 *
 * - 등록할 때 「시작 월」을 주면 그 달부터 이번 달까지 빠진 달을 한 번에 채운다 (api/expenses/recurring create).
 * - 매일 크론(api/cron/expense-mail)이 이번 달 분을 부른다 → 사장님이 버튼을 안 눌러도 매달 1일자로 생긴다.
 * - 멱등: 같은 달에 같은 분류·같은 금액 경비가 이미 있으면 건너뜀 (손으로 넣은 것·메일·카드로 확인한 것 포함).
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface RecurringItem { id: string; category: string; amount: number; memo: string | null; proof_type: string | null }

/** 한국시간 기준 이번 달 'YYYY-MM' */
export function thisMonthKst(): string {
  return new Date(Date.now() + 9 * 3600000).toISOString().slice(0, 7);
}

/** 'YYYY-MM' ~ 'YYYY-MM' 달 목록 (최대 36개월) */
export function monthsBetween(from: string, to: string): string[] {
  const out: string[] = [];
  let [y, m] = from.split('-').map(Number);
  const [ty, tm] = to.split('-').map(Number);
  while ((y < ty || (y === ty && m <= tm)) && out.length < 36) {
    out.push(`${y}-${String(m).padStart(2, '0')}`);
    m++; if (m > 12) { m = 1; y++; }
  }
  return out;
}

function monthEnd(month: string): string {
  const [y, m] = month.split('-').map(Number);
  return `${month}-${String(new Date(y, m, 0).getDate()).padStart(2, '0')}`;
}

export async function fillRecurring(db: Db, items: RecurringItem[], months: string[], userId?: string | null): Promise<{ generated: number; skipped: number }> {
  let generated = 0, skipped = 0;
  for (const item of items) {
    for (const month of months) {
      const { count } = await db.from('expenses').select('id', { count: 'exact', head: true })
        .gte('expense_date', `${month}-01`).lte('expense_date', monthEnd(month))
        .eq('category', item.category).eq('amount', item.amount);
      if ((count || 0) > 0) { skipped++; continue; }
      const { error } = await db.from('expenses').insert({
        expense_date: `${month}-01`,
        category: item.category,
        amount: item.amount,
        memo: `[고정] ${item.memo || item.category}`,
        kind: 'fixed',
        proof_type: item.proof_type || null,
        created_by: userId || null,
      });
      if (error) throw error;
      generated++;
    }
  }
  return { generated, skipped };
}
