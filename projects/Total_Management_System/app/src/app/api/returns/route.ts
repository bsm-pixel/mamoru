import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { insertReturn } from '@/lib/returns/insert-return';
import { toLocalDateString } from '@/lib/utils/format';

/** GET /api/returns?status=&search= — 반품·교환수거 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const sp = req.nextUrl.searchParams;
    const status = sp.get('status');
    const search = sp.get('search');

    let q = db.from('returns').select('*').order('created_at', { ascending: false }).limit(200);
    if (status && status !== 'all') q = q.eq('status', status);
    if (search) q = q.or(`return_number.ilike.%${search}%,name.ilike.%${search}%,phone.ilike.%${search}%,serial_number.ilike.%${search}%`);
    const { data, error } = await q;
    if (error) throw error;
    return NextResponse.json({ returns: data || [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/returns — 반품·교환수거 접수 생성 (교환 모달 '배송 수거' 분기에서 호출) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();

    const payload: Record<string, unknown> = {
      return_type: body.return_type || 'exchange',
      sale_id: body.sale_id || null,
      product_id: body.product_id || null,
      product_name: body.product_name || null,
      serial_id: body.serial_id || null,
      serial_number: body.serial_number || null,
      qty: body.qty ?? 1,
      customer_id: body.customer_id || null,
      name: body.name || null,
      phone: body.phone || null,
      postcode: body.postcode || null,
      address: body.address || null,
      address_detail: body.address_detail || null,
      pickup_method: body.pickup_method || null,
      pickup_date: body.pickup_date || null,
      status: 'requested',
      requested_at: new Date().toISOString(),
      refund_amount: body.refund_amount ?? 0,
      refund_method: body.refund_method || null,
      reason: body.reason || null,
      memo: body.memo || null,
      created_by: user.id,
    };

    const created = await insertReturn(db, toLocalDateString(new Date()), payload);
    return NextResponse.json({ return: created });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
