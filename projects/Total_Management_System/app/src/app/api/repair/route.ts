import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/repair — 복원수리 목록 (필터/페이징) */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const url = req.nextUrl;
    const status = url.searchParams.get('status');
    const search = url.searchParams.get('search');
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let query = (supabase as any)
      .from('repairs')
      .select('*', { count: 'exact' })
      .order('received_at', { ascending: false })
      .range(from, to);

    if (status && status !== 'all') {
      // 탭 그룹 필터 지원
      const statusGroups: Record<string, string[]> = {
        pickup: ['intake', 'pickup_scheduled', 'picked_up'],
        inspect: ['inspecting', 'cost_notified', 'payment_confirmed'],
        repair: ['repairing', 'ready_to_ship'],
        shipping: ['shipped', 'delivered'],
      };
      const group = statusGroups[status];
      if (group) {
        query = query.in('status', group);
      } else {
        query = query.eq('status', status);
      }
    }

    if (search) {
      query = query.or(
        `name.ilike.%${search}%,phone.ilike.%${search}%,as_id.ilike.%${search}%`
      );
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ repairs: data || [], total: count || 0 });
  } catch (err) {
    console.error('[repair] 목록 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/repair — 신규 접수 (TMS 직접 생성) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { data, error } = await db
      .from('repairs')
      .insert(body)
      .select()
      .single();

    if (error) throw error;

    // 초기 이력
    await db.from('repair_history').insert({
      repair_id: data.id,
      from_status: null,
      to_status: 'intake',
      changed_by: user.id,
      note: '신규 접수',
    });

    return NextResponse.json(data, { status: 201 });
  } catch (err) {
    console.error('[repair] 생성 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
