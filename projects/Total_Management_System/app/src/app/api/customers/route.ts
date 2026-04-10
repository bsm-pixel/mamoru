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
    const { name, phone, email, postcode, address_road, address_detail, customer_type, company_name, memo, tags } = body as {
      name: string;
      phone?: string;
      email?: string;
      postcode?: string;
      address_road?: string;
      address_detail?: string;
      customer_type?: string;
      company_name?: string;
      memo?: string;
      tags?: string[];
    };

    if (!name?.trim()) {
      return NextResponse.json({ error: '고객명은 필수입니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

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
