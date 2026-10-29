import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { fillRecurring, monthsBetween, thisMonthKst } from '@/lib/expenses/recurring';
import { toProof } from '@/lib/expenses/proof';

/** GET /api/expenses/recurring — 고정 경비 목록 */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (supabase as any)
      .from('recurring_expenses')
      .select('*')
      .eq('is_active', true)
      .order('created_at', { ascending: true });

    if (error) throw error;
    return NextResponse.json({ items: data || [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/expenses/recurring — 고정 경비 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();

    // action: 'create' | 'generate' | 'delete'
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    if (body.action === 'create') {
      const { category, amount, memo, from } = body;
      if (!amount || amount <= 0) return NextResponse.json({ error: '금액 필수' }, { status: 400 });

      const { data, error } = await db.from('recurring_expenses').insert({
        category: category || '기타',
        amount,
        memo: memo || null,
        proof_type: toProof(body.proof_type), // 169
      }).select().single();

      if (error) throw error;
      // 169 빠진 달 채우기 — 시작 월(YYYY-MM)부터 이번 달까지. 같은 달 같은 분류·금액이 있으면 건너뜀
      const now = thisMonthKst();
      const start = /^\d{4}-\d{2}$/.test(String(from || '')) && from <= now ? from : now;
      const filled = await fillRecurring(db, [data], monthsBetween(start, now), user.id);
      return NextResponse.json({ ...data, ...filled });
    }

    if (body.action === 'delete') {
      const { id } = body;
      await db.from('recurring_expenses').update({ is_active: false }).eq('id', id);
      return NextResponse.json({ ok: true });
    }

    if (body.action === 'generate') {
      // 이번달 고정 경비 — 169: 항목별로 이번 달 분만 (이미 있으면 건너뜀). 매일 크론도 같은 함수를 부른다
      const { data: items } = await db.from('recurring_expenses').select('id, category, amount, memo, proof_type').eq('is_active', true);
      if (!items || items.length === 0) {
        return NextResponse.json({ generated: 0, message: '등록된 고정 경비 없음' });
      }
      const res = await fillRecurring(db, items, [thisMonthKst()], user.id);
      return NextResponse.json({ generated: res.generated, message: res.generated ? undefined : '이번달 고정 경비가 이미 등록되어 있습니다' });
    }

    return NextResponse.json({ error: '알 수 없는 action' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
