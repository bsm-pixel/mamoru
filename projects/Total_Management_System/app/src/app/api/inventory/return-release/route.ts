import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { updateImwebStock } from '@/lib/imweb/client';

/**
 * 반품창고 검수 후 처리 (2026-10-14)
 *
 * 반품·교환으로 회수된 제품은 반품창고에 격리된다 (판매가능 현재고·아임웹 미반영).
 *   · 시리얼 제품   = product_serials status='returned', zone='return'
 *   · 비시리얼 제품 = products.return_stock
 * 사장님이 실물을 확인한 뒤 여기서 꺼낸다.
 *   restock   → 판매가능 재고로 복귀: 현재고 +1 + 아임웹 +1 (시리얼은 판매대기/진열 창고로 → 판매 입력 시리얼 목록에 뜸)
 *   defective → 불량 처리: 반품창고에서만 빠짐 (현재고·아임웹 변화 없음)
 * 무결성: 현재고 = 보관(raw_stock) + in_stock 시리얼 — 시리얼은 stock_quantity만 +1, 비시리얼은 raw_stock·stock_quantity 둘 다 +1.
 * 이력: stock_adjustments(재고 조정 이력) + 시리얼은 DB 트리거 audit log 자동.
 */

/** GET ?product_id= — 반품창고 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const productId = req.nextUrl.searchParams.get('product_id');
    if (!productId) return NextResponse.json({ error: 'product_id 가 필요합니다' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const [serialsRes, productRes] = await Promise.all([
      db.from('product_serials')
        .select('id, serial_number, updated_at')
        .eq('product_id', productId)
        .eq('status', 'returned')
        .order('updated_at', { ascending: false }),
      db.from('products').select('return_stock').eq('id', productId).single(),
    ]);
    if (serialsRes.error) throw serialsRes.error;

    return NextResponse.json({
      serials: serialsRes.data || [],
      // 마이그133 전이면 컬럼 없음 → 0
      return_stock: productRes.error ? 0 : (productRes.data?.return_stock || 0),
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST { product_id, serial_id?, action: 'restock'|'defective', zone?: 'ready'|'display' } */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { product_id, serial_id, action, zone } = await req.json() as {
      product_id?: string;
      serial_id?: string;
      action?: 'restock' | 'defective';
      zone?: 'ready' | 'display';
    };
    if (!product_id || (action !== 'restock' && action !== 'defective')) {
      return NextResponse.json({ error: '필수 값이 누락되었습니다' }, { status: 400 });
    }
    const targetZone = zone === 'display' ? 'display' : 'ready';

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: product } = await db
      .from('products')
      .select('id, name, stock_quantity, raw_stock, return_stock, imweb_product_no')
      .eq('id', product_id)
      .single();
    if (!product) return NextResponse.json({ error: '제품을 찾을 수 없습니다' }, { status: 404 });

    let label: string;

    if (serial_id) {
      // ── 시리얼 제품: returned 인 것만 (조건부 갱신 — 동시 클릭·중복 처리 방지)
      const { data: moved } = await db
        .from('product_serials')
        .update(action === 'restock'
          ? { status: 'in_stock', warehouse_zone: targetZone, previous_zone: 'return' }
          : { status: 'defective' })
        .eq('id', serial_id)
        .eq('product_id', product_id)
        .eq('status', 'returned')
        .select('serial_number')
        .maybeSingle();
      if (!moved) {
        return NextResponse.json({ error: '이미 처리됐거나 반품창고에 없는 시리얼입니다' }, { status: 409 });
      }
      label = `시리얼 ${moved.serial_number}`;

      if (action === 'restock') {
        // in_stock 시리얼 +1 → 현재고 +1 (보관 raw_stock 은 그대로)
        const { error } = await db.from('products')
          .update({ stock_quantity: (product.stock_quantity || 0) + 1, updated_at: new Date().toISOString() })
          .eq('id', product_id);
        if (error) throw error;
      }
    } else {
      // ── 비시리얼 제품: return_stock 1개
      const ret = product.return_stock || 0;
      if (ret < 1) return NextResponse.json({ error: '반품창고 재고가 없습니다' }, { status: 409 });
      const update: Record<string, unknown> = { return_stock: ret - 1, updated_at: new Date().toISOString() };
      if (action === 'restock') {
        update.raw_stock = (product.raw_stock || 0) + 1;
        update.stock_quantity = (product.stock_quantity || 0) + 1;
      }
      const { error } = await db.from('products').update(update).eq('id', product_id);
      if (error) throw error;
      label = '비시리얼 1개';
    }

    // 아임웹 재고 +1 (판매가능 복귀일 때만)
    let imwebSync: unknown = null;
    if (action === 'restock' && product.imweb_product_no) {
      try {
        imwebSync = await updateImwebStock(Number(product.imweb_product_no), 1);
      } catch (e) {
        imwebSync = { error: String(e) };
        console.error('[return-release] 아임웹 재고 동기화 실패:', product.imweb_product_no, e);
      }
    }

    // 재고 조정 이력 — 불량은 판매가능 수량 변화 0
    await db.from('stock_adjustments').insert({
      product_id,
      adjustment_type: action === 'restock' ? 'return' : 'damage',
      quantity: action === 'restock' ? 1 : 0,
      reason: action === 'restock'
        ? `반품창고 → 재판매 전환 (${label}${serial_id ? ` · ${targetZone === 'display' ? '진열' : '판매대기'}` : ''})`
        : `반품창고 → 불량 처리 (${label})`,
      adjusted_by: user.id,
    });

    return NextResponse.json({ ok: true, action, label, imweb_sync: imwebSync });
  } catch (err) {
    console.error('[return-release] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
