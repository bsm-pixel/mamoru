/**
 * 영수증 메일 → 경비 「확인 대기」 (2026-10-29)
 *
 * 매일 크론(api/cron/expense-mail) + 경비 화면 [메일 다시 확인] 버튼이 부른다.
 * bsm@mamoru.kr 메일함 하나만 읽는다 — mamoruscissors@gmail.com 등 다른 계정은 그쪽에서 자동 전달.
 *
 * - 같은 메일은 한 번만 (expense_inbox.gmail_id UNIQUE). 무시한 메일도 행이 남아 다시 안 올라온다.
 * - 거래처 기억(expense_vendor_rules)이 있으면 분류·고정/사용량을 그 값으로, auto_ignore 면 바로 ignored.
 * - 여기서는 expenses 에 아무것도 넣지 않는다. 확정은 사장님 [확인](api/expenses/inbox) 때만 — 회계 숫자는 사람이 확인한 것만.
 */
import { listMails } from '@/lib/google/gmail-client';
import { parseExpenseMail, expenseMailQuery } from './mail-parse';
import { estimateKrw } from './fx';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface ExpenseScanResult {
  ok: boolean;
  scanned: number;   // 새로 읽은 메일
  added: number;     // 확인 대기에 올린 수
  ignored: number;   // 거래처 규칙으로 바로 무시
  skipped: number;   // 영수증 아님
  needsReauth?: boolean;
  notConnected?: boolean;
  apiDisabled?: boolean;
  error?: string;
}

export async function runExpenseMailScan(db: Db, days = 3): Promise<ExpenseScanResult> {
  const base = { scanned: 0, added: 0, ignored: 0, skipped: 0 };

  // 이미 본 메일 id (최근 기간만 — 검색 범위보다 넉넉히)
  const since = new Date(Date.now() - (days + 2) * 86400000).toISOString();
  const { data: seenRows } = await db.from('expense_inbox').select('gmail_id').gte('received_at', since);
  const seen = new Set<string>((seenRows || []).map((r: { gmail_id: string }) => r.gmail_id));

  const res = await listMails(expenseMailQuery(days), days > 30 ? 300 : 100, seen);
  if (!res.ok) {
    return { ok: false, ...base, needsReauth: res.needsReauth, notConnected: res.notConnected, apiDisabled: res.apiDisabled, error: res.error };
  }

  const { data: ruleRows } = await db.from('expense_vendor_rules').select('vendor, category, kind, billing_cycle, auto_ignore');
  const rules = new Map<string, { category: string | null; kind: string | null; billing_cycle: string | null; auto_ignore: boolean }>(
    (ruleRows || []).map((r: { vendor: string; category: string | null; kind: string | null; billing_cycle: string | null; auto_ignore: boolean }) => [r.vendor, r]),
  );

  const out = { ...base };
  for (const m of res.messages) {
    if (seen.has(m.id)) continue;
    out.scanned++;
    const p = parseExpenseMail(m);
    if (!p) { out.skipped++; continue; }

    const rule = rules.get(p.vendor);
    const status = rule?.auto_ignore ? 'ignored' : 'pending';
    // 외화 영수증 — 결제일 환율로 원화 추정 (lib/expenses/fx.ts). 실패하면 빈칸 → 아래 fillFxEstimates 가 다음 스캔에 재시도
    if (p.amount === null && p.amountOriginal !== null && p.currency !== 'KRW') {
      p.amount = await estimateKrw(p.currency, p.amountOriginal, p.expenseDate);
    }
    const { error } = await db.from('expense_inbox').insert({
      gmail_id: m.id,
      received_at: m.receivedAt,
      sender: m.from.slice(0, 300),
      subject: m.subject.slice(0, 300),
      vendor: p.vendor,
      doc_type: p.docType,
      amount: p.amount,
      currency: p.currency,
      amount_original: p.amountOriginal,
      expense_date: p.expenseDate,
      category: rule?.category || p.category,
      kind: rule?.kind || p.kind,
      // 167 연간: 사장님이 고른 값(annual/monthly)이 우선, 기억이 없으면 메일 판독값
      billing_cycle: (rule?.billing_cycle === 'annual' || (!rule?.billing_cycle && p.billingCycle === 'annual')) && (rule?.kind || p.kind) === 'fixed' ? 'annual' : null,
      memo: p.memo || null,
      snippet: m.text.replace(/\s+/g, ' ').slice(0, 300),
      status,
      decided_at: status === 'ignored' ? new Date().toISOString() : null,
    });
    // 23505 = 동시에 두 번 돈 경우(크론·버튼) — 이미 들어간 것이므로 무시
    if (error && error.code !== '23505') throw error;
    if (!error) { if (status === 'ignored') out.ignored++; else out.added++; }
  }
  await fillFxEstimates(db);
  return { ok: true, ...out };
}

/** 확인 대기 중인 외화 영수증 중 원화가 빈 것 → 환율 추정으로 채움 (환율 실패분·기능 이전에 들어온 건 재시도) */
async function fillFxEstimates(db: Db): Promise<void> {
  const { data } = await db.from('expense_inbox')
    .select('id, currency, amount_original, expense_date, received_at')
    .eq('status', 'pending').is('amount', null).neq('currency', 'KRW').not('amount_original', 'is', null);
  for (const r of (data || []) as Array<{ id: string; currency: string; amount_original: number; expense_date: string | null; received_at: string | null }>) {
    const date = r.expense_date || (r.received_at || '').slice(0, 10);
    if (!date) continue;
    const krw = await estimateKrw(r.currency, Number(r.amount_original), date);
    // amount 가 아직 비어 있을 때만 (사장님이 그 사이 입력·확인했으면 건드리지 않음)
    if (krw !== null) await db.from('expense_inbox').update({ amount: krw }).eq('id', r.id).eq('status', 'pending').is('amount', null);
  }
}
