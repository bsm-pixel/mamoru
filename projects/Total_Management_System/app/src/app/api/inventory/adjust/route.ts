import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * POST /api/inventory/adjust — 재고 수동 조정
 * body: { product_id, adjustment_type, quantity, reason }
 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { product_id, adjustment_type, quantity, reason } = body;

    if (!product_id || !adjustment_type || quantity === undefined || quantity === 0) {
      return NextResponse.json({ error: '필수 항목을 입력해주세요 (제품, 유형, 수량)' }, { status: 400 });
    }

    const validTypes = ['damage', 'correction', 'return', 'other'];
    if (!validTypes.includes(adjustment_type)) {
      return NextResponse.json({ error: `유효하지 않은 조정 유형: ${adjustment_type}` }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 현재 재고 조회
    const { data: product, error: fetchErr } = await db
      .from('products')
      .select('id, name, stock_quantity')
      .eq('id', product_id)
      .single();

    if (fetchErr || !product) {
      return NextResponse.json({ error: '제품을 찾을 수 없습니다' }, { status: 404 });
    }

    const newQty = (product.stock_quantity || 0) + quantity;
    if (newQty < 0) {
      return NextResponse.json({
        error: `재고가 부족합니다. 현재 ${product.stock_quantity}개, 조정 ${quantity}개`,
      }, { status: 400 });
    }

    // 재고 업데이트
    const { error: updateErr } = await db
      .from('products')
      .update({ stock_quantity: newQty, updated_at: new Date().toISOString() })
      .eq('id', product_id);

    if (updateErr) throw updateErr;

    // 조정 이력 기록
    const { data: adjustment, error: insertErr } = await db
      .from('stock_adjustments')
      .insert({
        product_id,
        adjustment_type,
        quantity,
        reason: reason?.trim() || null,
        adjusted_by: user.id,
      })
      .select()
      .single();

    if (insertErr) throw insertErr;

    return NextResponse.json({
      adjustment,
      product_name: product.name,
      previous_qty: product.stock_quantity,
      new_qty: newQty,
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/**
 * GET /api/inventory/adjust — 조정 이력 조회
 * ?product_id=xxx&limit=20
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const sp = req.nextUrl.searchParams;
    const productId = sp.get('product_id');
    const limit = parseInt(sp.get('limit') || '20');

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    let query = db
      .from('stock_adjustments')
      .select('*, products(name, sku)')
      .order('created_at', { ascending: false })
      .limit(limit);

    if (productId) {
      query = query.eq('product_id', productId);
    }

    const { data, error } = await query;
    if (error) throw error;

    return NextResponse.json({ adjustments: data || [] });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
