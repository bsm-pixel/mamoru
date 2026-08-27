import { NextRequest, NextResponse } from 'next/server';
import { after } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { insertReturn } from '@/lib/returns/insert-return';
import { sendNotification } from '@/lib/notification/make-webhook';
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
      new_product_id: body.new_product_id || null,
      new_product_name: body.new_product_name || null,
      new_serial_number: body.new_serial_number || null,
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

    // 사장님 푸시(즉시) + 고객 알림톡(솔라피 return_received 등록 시). after()로 서버리스 완주 보장
    if (payload.phone) {
      after(async () => {
        try {
          await sendNotification({
            template: 'return_received',
            phone: String(payload.phone),
            name: String(payload.name || ''),
            data: {
              return_number: String(created.return_number || ''),
              product_name: String(payload.product_name || ''),
              pickup_method: String(payload.pickup_method || ''),
              pickup_date: String(payload.pickup_date || ''),
            },
          });
        } catch (e) { console.error('[returns POST] 알림 발송 실패(접수는 완료):', e); }
      });
    }

    return NextResponse.json({ return: created });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
