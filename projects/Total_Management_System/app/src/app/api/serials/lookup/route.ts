import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/serials/lookup?q=시리얼번호 — 시리얼 풀 정보 조회 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const q = req.nextUrl.searchParams.get('q')?.trim();
    if (!q || q.length < 2) {
      return NextResponse.json({ error: '검색어 2자 이상 필요' }, { status: 400 });
    }

    // 1. 시리얼 검색 (번호 or 바코드)
    let { data: serial } = await db
      .from('product_serials')
      .select('*')
      .eq('serial_number', q)
      .single();

    if (!serial) {
      const res = await db
        .from('product_serials')
        .select('*')
        .eq('barcode', q)
        .single();
      serial = res.data;
    }

    // 부분 매칭 (시리얼번호 포함)
    if (!serial) {
      const res = await db
        .from('product_serials')
        .select('*')
        .ilike('serial_number', `%${q}%`)
        .limit(1)
        .single();
      serial = res.data;
    }

    if (!serial) {
      return NextResponse.json({ serial: null, product: null, sale: null, repairs: [] });
    }

    // 2. 제품 정보
    const { data: product } = await db
      .from('products')
      .select('id, name, sku, category, image_url, price')
      .eq('id', serial.product_id)
      .single();

    // 3. 판매 정보 (sold일 때)
    let sale = null;
    if (serial.offline_sale_id) {
      const { data } = await db
        .from('offline_sales')
        .select('id, sale_number, sale_date, sale_channel, payment_method, payment_status, customer_name, customer_phone, total_amount')
        .eq('id', serial.offline_sale_id)
        .single();
      sale = data;
    }

    // 4. 복원수리 이력 (sold_to_phone 매칭)
    let repairs: Array<{ id: string; status: string; created_at: string }> = [];
    if (serial.sold_to_phone) {
      const { data } = await db
        .from('repairs')
        .select('id, status, created_at')
        .eq('customer_phone', serial.sold_to_phone)
        .order('created_at', { ascending: false })
        .limit(5);
      repairs = data || [];
    }

    return NextResponse.json({ serial, product, sale, repairs });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : JSON.stringify(err);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
