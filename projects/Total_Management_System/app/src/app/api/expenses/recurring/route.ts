import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

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
      const { category, amount, memo } = body;
      if (!amount || amount <= 0) return NextResponse.json({ error: '금액 필수' }, { status: 400 });

      const { data, error } = await db.from('recurring_expenses').insert({
        category: category || '기타',
        amount,
        memo: memo || null,
      }).select().single();

      if (error) throw error;
      return NextResponse.json(data);
    }

    if (body.action === 'delete') {
      const { id } = body;
      await db.from('recurring_expenses').update({ is_active: false }).eq('id', id);
      return NextResponse.json({ ok: true });
    }

    if (body.action === 'generate') {
      // 이번달 고정 경비 일괄 생성
      const { data: items } = await db
        .from('recurring_expenses')
        .select('category, amount, memo')
        .eq('is_active', true);

      if (!items || items.length === 0) {
        return NextResponse.json({ generated: 0, message: '등록된 고정 경비 없음' });
      }

      const now = new Date();
      const monthDate = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-01`;

      // 이미 이번달에 생성됐는지 체크 (중복 방지)
      const { count } = await db
        .from('expenses')
        .select('id', { count: 'exact', head: true })
        .eq('expense_date', monthDate)
        .ilike('memo', '%[고정]%');

      if ((count || 0) > 0) {
        return NextResponse.json({ generated: 0, message: '이번달 고정 경비가 이미 등록되어 있습니다' });
      }

      const rows = items.map((item: { category: string; amount: number; memo: string | null }) => ({
        expense_date: monthDate,
        category: item.category,
        amount: item.amount,
        memo: `[고정] ${item.memo || item.category}`,
        created_by: user.id,
      }));

      await db.from('expenses').insert(rows);
      return NextResponse.json({ generated: rows.length });
    }

    return NextResponse.json({ error: '알 수 없는 action' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
