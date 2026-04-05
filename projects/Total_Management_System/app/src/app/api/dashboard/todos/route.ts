import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/dashboard/todos — 할일 목록 (최신순) */
export async function GET() {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { data, error } = await (supabase as any)
      .from('dashboard_todos')
      .select('*')
      .eq('completed', false)
      .order('created_at', { ascending: false })
      .limit(20);
    if (error) throw error;
    return NextResponse.json({ todos: data || [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/dashboard/todos — 할일 추가 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { text } = await req.json();
    if (!text?.trim()) return NextResponse.json({ error: 'text required' }, { status: 400 });

    const { data, error } = await (supabase as any)
      .from('dashboard_todos')
      .insert({ text: text.trim() })
      .select()
      .single();
    if (error) throw error;
    return NextResponse.json({ todo: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/dashboard/todos — 할일 삭제 (완료 처리) */
export async function DELETE(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { id } = await req.json();
    if (!id) return NextResponse.json({ error: 'id required' }, { status: 400 });

    const { error } = await (supabase as any)
      .from('dashboard_todos')
      .delete()
      .eq('id', id);
    if (error) throw error;
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
