import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/customers — 고객 목록 (검색/필터/페이징) */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const url = req.nextUrl;
    const search = url.searchParams.get('search');
    const customerType = url.searchParams.get('type');
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    let query = db
      .from('customers')
      .select('*', { count: 'exact' })
      .is('merged_into_id', null) // 병합으로 흡수된 고객 숨김
      .order('created_at', { ascending: false })
      .range(from, to);

    if (search) {
      const digits = search.replace(/\D/g, '');
      const hasDigits = digits.length >= 2;
      const orFilter = hasDigits
        ? `name.ilike.%${search}%,phone_normalized.ilike.%${digits}%,company_name.ilike.%${search}%`
        : `name.ilike.%${search}%,company_name.ilike.%${search}%`;
      query = query.or(orFilter);
    }

    if (customerType) {
      query = query.eq('customer_type', customerType);
    } else {
      query = query.neq('customer_type', 'supplier'); // 고객 목록에서 매입처 제외
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ customers: data || [], total: count || 0 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/customers — 고객 신규 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { name, phone, email, postcode, address_road, address_detail, customer_type, company_name, activity_name, position, memo, tags, force } = body as {
      name: string;
      phone?: string;
      email?: string;
      postcode?: string;
      address_road?: string;
      address_detail?: string;
      customer_type?: string;
      company_name?: string;
      activity_name?: string;
      position?: string;
      memo?: string;
      tags?: string[];
      force?: boolean; // 같은 전화 중복 경고 무시하고 강제 등록
    };

    if (!name?.trim()) {
      return NextResponse.json({ error: '고객명은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 같은 전화번호 중복 방지 — 수기 등록도 phone_normalized 기준으로 기존 고객 검사 (접수 자동매칭과 정합)
    const phoneNorm = (phone || '').replace(/\D/g, '');
    if (phoneNorm && !force) {
      const { data: dup } = await db
        .from('customers')
        .select('id, name, phone, customer_type, company_name, source')
        .eq('phone_normalized', phoneNorm)
        .is('merged_into_id', null)
        .limit(1);
      if (dup && dup.length > 0) {
        return NextResponse.json(
          { error: '이미 등록된 전화번호입니다', existing: dup[0] },
          { status: 409 }
        );
      }
    }

    const { data: customer, error } = await db
      .from('customers')
      .insert({
        name: name.trim(),
        phone: phone || null,
        email: email || null,
        postcode: postcode || null,
        address_road: address_road || null,
        address_detail: address_detail || null,
        customer_type: customer_type || 'retail',
        company_name: company_name || null,
        activity_name: activity_name?.trim() || null,
        position: position?.trim() || null,
        memo: memo || null,
        tags: Array.isArray(tags) ? tags : [],
        source: 'manual',
      })
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json({ customer });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
