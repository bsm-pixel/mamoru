import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { randomBytes } from 'crypto';

/** POST /api/serials — 시리얼 단건 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();

    // raw_stock 체크 — 보관창고에 재고 있어야 시리얼 생성 가능
    const { data: product } = await db
      .from('products')
      .select('raw_stock')
      .eq('id', body.product_id)
      .single();

    if (!product || (product.raw_stock || 0) < 1) {
      return NextResponse.json({ error: '보관창고 재고가 부족합니다' }, { status: 400 });
    }

    const { data, error } = await db
      .from('product_serials')
      .insert({
        product_id: body.product_id,
        serial_number: body.serial_number,
        barcode: body.barcode || body.serial_number,
        verify_token: randomBytes(6).toString('hex'),
        warehouse_zone: 'ready', // 시리얼 생성 = 마킹 완료 → 준비 창고
        lot_number: body.lot_number || null,
        manufactured_at: body.manufactured_at || null,
        memo: body.memo || null,
        created_by: user.id,
      })
      .select()
      .single();

    if (error) throw error;

    // raw_stock 차감 (보관에서 꺼냄, stock_quantity는 유지 — 총합 동일)
    await db
      .from('products')
      .update({ raw_stock: (product.raw_stock || 0) - 1 })
      .eq('id', body.product_id);

    return NextResponse.json({ serial: data });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH /api/serials — 일괄 warehouse_zone 변경 */
export async function PATCH(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { ids, warehouse_zone } = await req.json() as {
      ids: string[];
      warehouse_zone: 'raw' | 'ready' | 'display';
    };

    if (!ids?.length || !warehouse_zone) {
      return NextResponse.json({ error: 'ids와 warehouse_zone 필수' }, { status: 400 });
    }

    const { error } = await db
      .from('product_serials')
      .update({ warehouse_zone })
      .in('id', ids);

    if (error) throw error;
    return NextResponse.json({ success: true, updated: ids.length });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
