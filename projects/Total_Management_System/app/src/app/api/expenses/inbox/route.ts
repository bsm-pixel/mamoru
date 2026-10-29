import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { runExpenseMailScan } from '@/lib/expenses/mail-scan';

/**
 * 메일에서 찾은 영수증 — 확인 대기함 (2026-10-29, 마이그 166)
 *
 * GET  /api/expenses/inbox            → 확인 대기 목록 (+ 같은 금액·가까운 날짜의 기존 경비 = 중복 의심)
 * POST { action:'confirm', id, amount, expense_date, category, kind, memo }
 *        → expenses 1건 생성 + 거래처 기억(분류·고정/사용량) 저장
 * POST { action:'ignore', id, always? }   → 무시 (always=true 면 그 거래처는 다음부터 자동 무시)
 * POST { action:'scan', days? }           → 지금 메일 다시 확인 (처음엔 days=120 으로 과거분)
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
  memo: string | null;
  dup_hint?: string | null; // 이미 같은 금액 경비가 근처 날짜에 있으면 그 설명
}

export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const db = supabase as Db;

    const { data, error } = await db.from('expense_inbox')
      .select('id, gmail_id, received_at, sender, subject, vendor, doc_type, amount, currency, amount_original, expense_date, category, kind, memo')
      .eq('status', 'pending')
      .order('received_at', { ascending: false })
      .limit(200);
    if (error) throw error;
    const items = (data || []) as ExpenseInboxItem[];

    // 중복 의심: 같은 금액 경비가 ±7일 안에 이미 있음 (손으로 넣은 것 + 카드 영수증·세금계산서 이중 등)
    const amounts = [...new Set(items.map((i) => i.amount).filter((a): a is number => !!a))];
    if (amounts.length) {
      const { data: ex } = await db.from('expenses').select('expense_date, amount, memo, category').in('amount', amounts);
      for (const it of items) {
        if (!it.amount || !it.expense_date) continue;
        const t = new Date(it.expense_date).getTime();
        const hit = (ex || []).find((e: { expense_date: string; amount: number }) =>
          e.amount === it.amount && Math.abs(new Date(e.expense_date).getTime() - t) <= 7 * 86400000);
        if (hit) it.dup_hint = `${hit.expense_date} ${hit.category}${hit.memo ? ` · ${hit.memo}` : ''}`;
      }
    }
    return NextResponse.json({ items });
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

    const { id } = body;
    if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });
    const { data: item } = await db.from('expense_inbox').select('*').eq('id', id).single();
    if (!item) return NextResponse.json({ error: '항목을 찾을 수 없습니다' }, { status: 404 });
    if (item.status !== 'pending') return NextResponse.json({ error: '이미 처리된 항목입니다' }, { status: 409 });
    const now = new Date().toISOString();

    if (body.action === 'ignore') {
      // CAS — 두 번 눌러도 한 번만
      const { data: done } = await db.from('expense_inbox')
        .update({ status: 'ignored', decided_at: now }).eq('id', id).eq('status', 'pending').select('id');
      if (!done?.length) return NextResponse.json({ error: '이미 처리된 항목입니다' }, { status: 409 });
      if (body.always && item.vendor) {
        await db.from('expense_vendor_rules').upsert({ vendor: item.vendor, auto_ignore: true, updated_at: now });
      }
      return NextResponse.json({ ok: true });
    }

    if (body.action === 'confirm') {
      const amount = parseInt(String(body.amount ?? item.amount ?? ''), 10);
      if (!amount) return NextResponse.json({ error: '금액을 입력해 주세요' }, { status: 400 });
      const kind = body.kind === 'fixed' ? 'fixed' : 'variable';
      const category = String(body.category || item.category || '기타');

      // 먼저 대기함을 선점(CAS) → 경비 생성. 생성 실패 시 대기로 되돌린다(이중 생성 방지)
      const { data: claimed } = await db.from('expense_inbox')
        .update({ status: 'confirmed', decided_at: now }).eq('id', id).eq('status', 'pending').select('id');
      if (!claimed?.length) return NextResponse.json({ error: '이미 처리된 항목입니다' }, { status: 409 });

      // 외화: 환율 추정액(대기함 금액)을 그대로 확인하면 "환율 추정" 표시 — 명세서 금액으로 고쳤으면 원금만 남김
      const fxEst = item.currency !== 'KRW' && item.amount === amount ? ' · 환율 추정' : '';
      const usd = item.currency !== 'KRW' && item.amount_original ? ` (${item.currency} ${item.amount_original}${fxEst})` : '';
      const { data: expense, error: insErr } = await db.from('expenses').insert({
        expense_date: body.expense_date || item.expense_date || now.slice(0, 10),
        category,
        amount,
        kind,
        vendor: item.vendor,
        memo: (String(body.memo ?? '').trim() || [item.vendor, item.memo].filter(Boolean).join(' · ')) + usd || null,
        source_ref: `mail:${item.gmail_id}`,
        created_by: user.id,
      }).select('id').single();
      if (insErr) {
        await db.from('expense_inbox').update({ status: 'pending', decided_at: null }).eq('id', id);
        throw insErr;
      }
      await db.from('expense_inbox').update({ expense_id: expense.id, amount, category, kind }).eq('id', id);
      // 거래처 기억 — 다음 메일부터 같은 분류·고정/사용량으로 올라온다
      if (item.vendor) {
        await db.from('expense_vendor_rules').upsert({ vendor: item.vendor, category, kind, auto_ignore: false, updated_at: now });
      }
      return NextResponse.json({ ok: true, expense_id: expense.id });
    }

    return NextResponse.json({ error: 'unknown action' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
