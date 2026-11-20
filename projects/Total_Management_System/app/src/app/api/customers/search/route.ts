import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { NOT_HIDDEN_FILTER } from '@/lib/customer/hidden';
import { customerSearchOr } from '@/lib/customer/search-filter';

/** GET /api/customers/search?q=홍 — 고객 자동완성 검색 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // 2026-11-20 검색 조건은 한 곳(search-filter)에서 — 고객 목록과 같은 칸을 찾는다(활동명 · 상세 주소 포함)
    const orFilter = customerSearchOr(req.nextUrl.searchParams.get('q'));
    if (!orFilter) {
      return NextResponse.json({ customers: [] });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data, error } = await db
      .from('customers')
      .select('id, name, phone, email, address_road, address_detail, postcode, ecount_customer_code, source, customer_type, company_name, memo, default_repair_price')
      .or(orFilter)
      .neq('customer_type', 'supplier') // 판매 자동완성에서 매입처 제외
      .is('merged_into_id', null) // 병합으로 흡수된 고객 숨김
      .or(NOT_HIDDEN_FILTER) // 2026-11-15 숨긴 고객(꼬리표 「숨김」) 제외
      .order('name')
      .limit(10);

    if (error) throw error;

    return NextResponse.json({ customers: data || [] });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
