import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

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
    const byKind = { fixed: 0, variable: 0, none: 0 };
    let total = 0;
    for (const e of (data || [])) {
      byCategory[e.category] = (byCategory[e.category] || 0) + e.amount;
      if (e.kind === 'fixed') byKind.fixed += e.amount;
      else if (e.kind === 'variable') byKind.variable += e.amount;
      else byKind.none += e.amount;
      total += e.amount;
    }

    return NextResponse.json({ expenses: data || [], total, byCategory, byKind });
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
    const { expense_date, category, amount, memo, kind } = body;

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
    const { id, expense_date, category, amount, memo, kind } = body;
    if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });

    const updateData: Record<string, unknown> = {};
    if (expense_date !== undefined) updateData.expense_date = expense_date;
    if (category !== undefined) updateData.category = category;
    if (amount !== undefined) updateData.amount = amount;
    if (memo !== undefined) updateData.memo = memo || null;
    if (kind !== undefined) updateData.kind = kind === 'fixed' || kind === 'variable' ? kind : null; // 166

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { error } = await (supabase as any).from('expenses').update(updateData).eq('id', id);
    if (error) throw error;
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
