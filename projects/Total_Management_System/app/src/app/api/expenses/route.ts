import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { toProof } from '@/lib/expenses/proof';

/** GET /api/expenses — 경비 목록 (기간 필터) */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const sp = req.nextUrl.searchParams;
    const from = sp.get('from');
    const to = sp.get('to');

    let query = db.from('expenses').select('*').order('expense_date', { ascending: false });
    if (from) query = query.gte('expense_date', from);
    if (to) query = query.lte('expense_date', to);

    const { data, error } = await query;
    if (error) throw error;

    // 카테고리별 합계
    const byCategory: Record<string, number> = {};
    // 166: 고정(구독·임대) / 사용량(택배·부자재) 합계 — kind 없는 옛 데이터는 none(미분류)
    // 167: annual = 그중 연간 결제(고정에 포함된 금액 — 장부는 결제일 전액, 현금주의 그대로)
    const byKind = { fixed: 0, variable: 0, none: 0, annual: 0 };
    let total = 0;
    for (const e of (data || [])) {
      byCategory[e.category] = (byCategory[e.category] || 0) + e.amount;
      if (e.kind === 'fixed') {
        byKind.fixed += e.amount;
        if (e.billing_cycle === 'annual') byKind.annual += e.amount;
      }
      else if (e.kind === 'variable') byKind.variable += e.amount;
      else byKind.none += e.amount;
      total += e.amount;
    }

    // 167: 월 환산 고정비 = 기간 안 매달 고정비 + (기간 끝 기준 최근 1년 연간 결제 ÷ 12 × 기간 개월 수)
    //   연간 결제가 결제한 달에만 크게 잡혀 "한 달 고정비"가 들쭉날쭉해 보이는 것을 보정한 참고값. 장부·리포트엔 안 쓴다.
    const endDate = to || new Date().toISOString().slice(0, 10);
    const startDate = from || endDate;
    const months = Math.max(1, Math.round(((Date.parse(endDate) - Date.parse(startDate)) / 86400000 + 1) / 30.44));
    const yearAgo = new Date(Date.parse(endDate) - 365 * 86400000).toISOString().slice(0, 10);
    const { data: annualRows } = await db.from('expenses').select('amount')
      .eq('kind', 'fixed').eq('billing_cycle', 'annual').gt('expense_date', yearAgo).lte('expense_date', endDate);
    const annualYear = (annualRows || []).reduce((s: number, r: { amount: number }) => s + r.amount, 0);
    const fixedMonthly = Math.round(byKind.fixed - byKind.annual + (annualYear / 12) * months);

    return NextResponse.json({ expenses: data || [], total, byCategory, byKind, fixedMonthly, fixedMonths: months });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/expenses — 경비 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { expense_date, category, amount, memo, kind, billing_cycle, proof_type } = body;

    if (!amount || amount <= 0) {
      return NextResponse.json({ error: '금액은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (supabase as any)
      .from('expenses')
      .insert({
        expense_date: expense_date || new Date().toISOString().slice(0, 10),
        category: category || '기타',
        amount,
        memo: memo || null,
        kind: kind === 'fixed' || kind === 'variable' ? kind : null, // 166 고정/사용량
        billing_cycle: kind === 'fixed' && billing_cycle === 'annual' ? 'annual' : null, // 167 연간
        proof_type: toProof(proof_type), // 168 증빙
        created_by: user.id,
      })
      .select()
      .single();

    if (error) throw error;
    return NextResponse.json(data);
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/expenses?id=xxx — 경비 삭제 */
export async function DELETE(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const id = req.nextUrl.searchParams.get('id');
    if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { error } = await (supabase as any).from('expenses').delete().eq('id', id);
    if (error) throw error;
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/expenses — 경비 수정 */
export async function PATCH(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { id, expense_date, category, amount, memo, kind, billing_cycle, proof_type } = body;
    if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });

    const updateData: Record<string, unknown> = {};
    if (expense_date !== undefined) updateData.expense_date = expense_date;
    if (category !== undefined) updateData.category = category;
    if (amount !== undefined) updateData.amount = amount;
    if (memo !== undefined) updateData.memo = memo || null;
    if (kind !== undefined) updateData.kind = kind === 'fixed' || kind === 'variable' ? kind : null; // 166
    // 167 연간: billing_cycle 을 보낼 때만 바꾼다(kind 만 바꿀 때 기존 연간 표시 유지). 사용량으로 바꾸면 연간 해제
    if (billing_cycle !== undefined) updateData.billing_cycle = billing_cycle === 'annual' && kind !== 'variable' ? 'annual' : null;
    else if (kind === 'variable') updateData.billing_cycle = null;
    if (proof_type !== undefined) updateData.proof_type = toProof(proof_type); // 168 증빙 (null = 미정으로)

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { error } = await (supabase as any).from('expenses').update(updateData).eq('id', id);
    if (error) throw error;
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
