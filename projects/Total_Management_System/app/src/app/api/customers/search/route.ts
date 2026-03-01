import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/customers/search?q=홍 — 고객 자동완성 검색 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const q = req.nextUrl.searchParams.get('q')?.trim();
    if (!q || q.length < 2) {
      return NextResponse.json({ customers: [] });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 숫자만 추출 (전화번호 검색용)
    const digits = q.replace(/\D/g, '');
    const hasDigits = digits.length >= 2;

    // 이름 OR 전화번호(정규화)로 부분매칭 검색
    const orFilter = hasDigits
      ? `name.ilike.%${q}%,phone_normalized.ilike.%${digits}%`
      : `name.ilike.%${q}%`;

    const { data, error } = await db
      .from('customers')
      .select('id, name, phone, email, address_road, address_detail, postcode, ecount_customer_code, source')
      .or(orFilter)
      .order('name')
      .limit(10);

    if (error) throw error;

    return NextResponse.json({ customers: data || [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
