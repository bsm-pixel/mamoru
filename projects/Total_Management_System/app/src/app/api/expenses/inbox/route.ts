import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { runExpenseMailScan } from '@/lib/expenses/mail-scan';
import { toProof } from '@/lib/expenses/proof';

/**
 * 메일·카드 내역에서 찾은 경비 — 확인 대기함 (2026-10-29, 마이그 166)
 *
 * GET  /api/expenses/inbox            → 확인 대기 목록 (+ 중복 의심 dup_hint)
 * POST { action:'confirm', id, amount, expense_date, category, kind, billing_cycle(167 'annual'|null), proof_type(168), memo }
 *        → expenses 1건 생성 + 거래처 기억(분류·고정/사용량·증빙) 저장
 * POST { action:'confirm_known' }          → 거래처 기억이 있고 중복 의심이 없는 대기 건을 기억값 그대로 한 번에 확인
 * POST { action:'ignore', id, always? }   → 무시 (always=true 면 그 거래처는 다음부터 자동 무시)
 * POST { action:'scan', days? }           → 지금 메일 다시 확인 (처음엔 days=120 으로 과거분)
 * 카드 이용내역 엑셀은 api/expenses/card-import 가 같은 대기함에 넣는다 (gmail_id = 'card:…').
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface ExpenseInboxItem {
  id: string;
  gmail_id: string;
  received_at: string | null;
  sender: string | null;
  subject: string | null;
  vendor: string | null;
  doc_type: 'card' | 'receipt' | 'invoice' | 'tax_invoice';
  amount: number | null;
  currency: string;
  amount_original: number | null;
  expense_date: string | null;
  category: string | null;
  kind: 'fixed' | 'variable' | null;
  billing_cycle: 'annual' | 'monthly' | null; // 167 연간 구독
  proof_type: string | null; // 168 증빙 (lib/expenses/proof)
  memo: string | null;
  dup_hint?: string | null; // 이미 같은 경비가 근처 날짜에 있으면 그 설명
}

const DAY = 86400000;
/** 거래처 비교용 — 첫 단어 소문자 (「ANTHROPIC」 카드 줄 ↔ 「Anthropic」 메일 영수증) */
const vkey = (v: string | null | undefined) => (v || '').toLowerCase().replace(/[^a-z0-9가-힣 ]/g, ' ').trim().split(/\s+/)[0] || '';

/**
 * 중복 의심 표시:
 *  ① 이미 등록된 경비 — 같은 금액 ±7일, 또는 같은 거래처 ±5일(외화는 환율 추정과 카드 청구액이 달라 금액이 안 맞음)
 *  ② 카드 내역 줄 — 대기함의 메일 건과 같은 금액 ±3일 (카드 승인 메일 + 카드 엑셀 이중)
 */
async function markDuplicates(db: Db, items: ExpenseInboxItem[]): Promise<void> {
  const dates = items.map((i) => i.expense_date).filter((d): d is string => !!d).sort();
  if (!dates.length) return;
  const lo = new Date(new Date(dates[0]).getTime() - 7 * DAY).toISOString().slice(0, 10);
  const hi = new Date(new Date(dates[dates.length - 1]).getTime() + 7 * DAY).toISOString().slice(0, 10);
  const { data: ex } = await db.from('expenses').select('expense_date, amount, memo, category, vendor')
    .gte('expense_date', lo).lte('expense_date', hi);
  for (const it of items) {
    if (!it.expense_date) continue;
    const t = new Date(it.expense_date).getTime();
    const k = vkey(it.vendor);
    const hit = (ex || []).find((e: { expense_date: string; amount: number; vendor: string | null }) => {
      const gap = Math.abs(new Date(e.expense_date).getTime() - t);
      return (!!it.amount && e.amount === it.amount && gap <= 7 * DAY) || (!!k && k.length >= 3 && vkey(e.vendor) === k && gap <= 5 * DAY);
    });
    if (hit) { it.dup_hint = `${hit.expense_date} ${hit.category}${hit.memo ? ` · ${hit.memo}` : ''}`; continue; }
    if (it.gmail_id.startsWith('card:') && it.amount) {
      const twin = items.find((o) => !o.gmail_id.startsWith('card:') && o.amount === it.amount && o.expense_date
        && Math.abs(new Date(o.expense_date).getTime() - t) <= 3 * DAY);
      if (twin) it.dup_hint = `대기함 메일 「${twin.vendor}」와 같은 금액 — 둘 중 하나만 확인`;
    }
  }
}

async function loadPending(db: Db): Promise<ExpenseInboxItem[]> {
  const { data, error } = await db.from('expense_inbox')
    .select('id, gmail_id, received_at, sender, subject, vendor, doc_type, amount, currency, amount_original, expense_date, category, kind, billing_cycle, proof_type, memo')
    .eq('status', 'pending')
    .order('received_at', { ascending: false })
    .limit(1000); // 카드 내역 1년치(수백 건)도 한 번에
  if (error) throw error;
  const items = (data || []) as ExpenseInboxItem[];
  await markDuplicates(db, items);
  return items;
}

type ConfirmInput = { amount?: unknown; expense_date?: unknown; category?: unknown; kind?: unknown; billing_cycle?: unknown; proof_type?: unknown; memo?: unknown };

/** 대기 1건 → 경비. 409/400 은 { error, status } 로 돌려준다 */
async function confirmItem(db: Db, item: Record<string, unknown> & ExpenseInboxItem & { status: string }, body: ConfirmInput, userId: string): Promise<{ expense_id?: string; error?: string; status?: number }> {
  if (item.status !== 'pending') return { error: '이미 처리된 항목입니다', status: 409 };
  const amount = parseInt(String(body.amount ?? item.amount ?? ''), 10);
  if (!amount) return { error: '금액을 입력해 주세요', status: 400 };
  const kind = (body.kind ?? item.kind) === 'fixed' ? 'fixed' : 'variable';
  const category = String(body.category || item.category || '기타');
  // 167 연간 — 고정일 때만. 화면 [연간] 선택 = billing_cycle 'annual'
  const cycle = kind === 'fixed' && (body.billing_cycle ?? item.billing_cycle) === 'annual' ? 'annual' : null;
  const proof = toProof(body.proof_type) ?? toProof(item.proof_type); // 168
  const now = new Date().toISOString();

  // 먼저 대기함을 선점(CAS) → 경비 생성. 생성 실패 시 대기로 되돌린다(이중 생성 방지)
  const { data: claimed } = await db.from('expense_inbox')
    .update({ status: 'confirmed', decided_at: now }).eq('id', item.id).eq('status', 'pending').select('id');
  if (!claimed?.length) return { error: '이미 처리된 항목입니다', status: 409 };

  // 외화: 환율 추정액(대기함 금액)을 그대로 확인하면 "환율 추정" 표시 — 명세서 금액으로 고쳤으면 원금만 남김
  const fxEst = item.currency !== 'KRW' && item.amount === amount ? ' · 환율 추정' : '';
  const usd = item.currency !== 'KRW' && item.amount_original ? ` (${item.currency} ${item.amount_original}${fxEst})` : '';
  const isCard = item.gmail_id.startsWith('card:');
  const { data: expense, error: insErr } = await db.from('expenses').insert({
    expense_date: body.expense_date || item.expense_date || now.slice(0, 10),
    category,
    amount,
    kind,
    billing_cycle: cycle,
    proof_type: proof,
    vendor: item.vendor,
    memo: (String(body.memo ?? '').trim() || [item.vendor, item.memo].filter(Boolean).join(' · ')) + usd || null,
    source_ref: isCard ? item.gmail_id : `mail:${item.gmail_id}`, // 카드 내역 = 'card:…'
    created_by: userId,
  }).select('id').single();
  if (insErr) {
    await db.from('expense_inbox').update({ status: 'pending', decided_at: null }).eq('id', item.id);
    throw insErr;
  }
  await db.from('expense_inbox').update({ expense_id: expense.id, amount, category, kind, billing_cycle: cycle, proof_type: proof }).eq('id', item.id);
  // 거래처 기억 — 다음 메일·카드 줄부터 같은 분류·고정/사용량·증빙으로 올라온다
  if (item.vendor) {
    await db.from('expense_vendor_rules').upsert({ vendor: item.vendor, category, kind, billing_cycle: kind === 'fixed' ? (cycle || 'monthly') : null, proof_type: proof, auto_ignore: false, updated_at: now });
  }
  return { expense_id: expense.id };
}

export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const items = await loadPending(supabase as Db);
    // 거래처 기억이 있는 건 수 — 화면 [기억된 거래처 한 번에 확인] 버튼용
    const { data: rules } = await (supabase as Db).from('expense_vendor_rules').select('vendor').eq('auto_ignore', false);
    const known = new Set((rules || []).map((r: { vendor: string }) => r.vendor));
    const knownCount = items.filter((i) => i.vendor && known.has(i.vendor) && i.amount && !i.dup_hint).length;
    return NextResponse.json({ items, knownCount });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const db = supabase as Db;
    const body = await req.json();

    if (body.action === 'scan') {
      const days = Math.min(365, Math.max(1, parseInt(String(body.days || 3), 10) || 3));
      // 구글 토큰·전체 행 접근이 필요해 서비스 클라이언트로 (크론과 같은 경로)
      const result = await runExpenseMailScan(createServiceClient() as Db, days);
      return NextResponse.json(result);
    }

    if (body.action === 'confirm_known') {
      // 기억된 거래처 + 금액 있음 + 중복 의심 없음 → 기억값(대기함에 이미 채워진 분류·구분·증빙) 그대로 확인
      const items = await loadPending(db);
      const { data: rules } = await db.from('expense_vendor_rules').select('vendor').eq('auto_ignore', false);
      const known = new Set((rules || []).map((r: { vendor: string }) => r.vendor));
      let confirmed = 0;
      for (const it of items) {
        if (!it.vendor || !known.has(it.vendor) || !it.amount || it.dup_hint) continue;
        const r = await confirmItem(db, { ...it, status: 'pending' }, {}, user.id);
        if (r.expense_id) confirmed++;
      }
      return NextResponse.json({ ok: true, confirmed });
    }

    const { id } = body;
    if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });
    const { data: item } = await db.from('expense_inbox').select('*').eq('id', id).single();
    if (!item) return NextResponse.json({ error: '항목을 찾을 수 없습니다' }, { status: 404 });
    if (item.status !== 'pending') return NextResponse.json({ error: '이미 처리된 항목입니다' }, { status: 409 });

    if (body.action === 'ignore') {
      // CAS — 두 번 눌러도 한 번만
      const { data: done } = await db.from('expense_inbox')
        .update({ status: 'ignored', decided_at: new Date().toISOString() }).eq('id', id).eq('status', 'pending').select('id');
      if (!done?.length) return NextResponse.json({ error: '이미 처리된 항목입니다' }, { status: 409 });
      if (body.always && item.vendor) {
        await db.from('expense_vendor_rules').upsert({ vendor: item.vendor, auto_ignore: true, updated_at: new Date().toISOString() });
        // 같은 거래처의 다른 대기 건도 같이 무시 (카드 내역에 같은 가게가 여러 줄)
        await db.from('expense_inbox').update({ status: 'ignored', decided_at: new Date().toISOString() })
          .eq('vendor', item.vendor).eq('status', 'pending');
      }
      return NextResponse.json({ ok: true });
    }

    if (body.action === 'confirm') {
      const r = await confirmItem(db, item, body, user.id);
      if (r.error) return NextResponse.json({ error: r.error }, { status: r.status || 400 });
      return NextResponse.json({ ok: true, expense_id: r.expense_id });
    }

    return NextResponse.json({ error: 'unknown action' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
