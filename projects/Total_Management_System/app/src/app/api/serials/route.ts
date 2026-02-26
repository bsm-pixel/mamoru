import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** POST /api/serials — 시리얼 단건 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();

    const { data, error } = await db
      .from('product_serials')
      .insert({
        product_id: body.product_id,
        serial_number: body.serial_number,
        barcode: body.barcode || null,
        lot_number: body.lot_number || null,
        manufactured_at: body.manufactured_at || null,
        memo: body.memo || null,
        created_by: user.id,
      })
      .select()
      .single();

    if (error) throw error;

    // stock_quantity 증가
    await db.rpc('increment_stock', { p_id: body.product_id, qty: 1 }).catch(() => {
      // RPC 없으면 수동 업데이트
      return db
        .from('products')
        .update({ stock_quantity: db.raw('stock_quantity + 1') })
        .eq('id', body.product_id);
    });

    return NextResponse.json({ serial: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
