import { NextRequest, NextResponse } from 'next/server';
import { bookSingle } from '@/lib/lotte/client';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/lotte/book — 송장 생성 */
export async function POST(request: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await request.json();
    const result = await bookSingle(body);

    // 주문에 송장 정보 저장
    if (body.orderId) {
      await (supabase as any)
        .from('orders')
        .update({
          invoice_number: result.invNo,
          courier_code: 'LOTTE',
          courier_name: '롯데택배',
          status: 'shipping',
          shipped_at: new Date().toISOString(),
        })
        .eq('id', body.orderId);
    }

    return NextResponse.json(result);
  } catch (err) {
    console.error('[lotte/book] 송장 생성 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
