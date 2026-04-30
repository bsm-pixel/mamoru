import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/consultation — 상담 목록 (필터/페이징) */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    const url = req.nextUrl;
    const status = url.searchParams.get('status');
    const type = url.searchParams.get('type');
    const search = url.searchParams.get('search');
    const phone = url.searchParams.get('phone'); // 070: 정규화된 phone 일치 검색
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let query = (supabase as any)
      .from('consultations')
      .select('*', { count: 'exact' })
      .order('received_at', { ascending: false })
      .range(from, to);

    if (status && status !== 'all') {
      query = query.eq('status', status);
    }
    if (type && type !== 'all') {
      query = query.eq('consultation_type', type);
    }
    if (phone) {
      // 070: phone_normalized 정확 매칭 (digits-only)
      const phoneDigits = phone.replace(/\D/g, '');
      if (phoneDigits) query = query.eq('phone_normalized', phoneDigits);
    }
    if (search) {
      query = query.or(
        `name.ilike.%${search}%,phone.ilike.%${search}%,unique_id.ilike.%${search}%,address_road.ilike.%${search}%`
      );
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ consultations: data || [], total: count || 0 });
  } catch (err) {
    console.error('[consultation] 목록 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/consultation — 신규 상담 생성 */
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
      .from('consultations')
      .insert(body)
      .select()
      .single();

    if (error) throw error;

    // 초기 이력 기록
    await db.from('consultation_history').insert({
      consultation_id: data.id,
      from_status: null,
      to_status: body.status || 'pending_admin',
      changed_by: user.id,
      note: '신규 접수',
    });

    return NextResponse.json(data, { status: 201 });
  } catch (err) {
    console.error('[consultation] 생성 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
