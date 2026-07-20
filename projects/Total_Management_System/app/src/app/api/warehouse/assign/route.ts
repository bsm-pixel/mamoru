import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/**
 * 제품 → 정위치 배정 API — 112, 2026-07-18
 *
 * PATCH /api/warehouse/assign   { product_ids: string[], location_id: string | null }
 *   location_id = null 이면 위치 해제(미지정)
 *
 * ⚠️ products 테이블에서 location_id 만 건드린다.
 *    stock_quantity / raw_stock 는 읽지도 쓰지도 않는다 — 재고 정합성 무영향 보장.
 */
export async function PATCH(req: NextRequest) {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  try {
    const body = await req.json();
    const productIds: string[] = Array.isArray(body.product_ids) ? body.product_ids : [];
    const locationId: string | null = body.location_id ?? null;

    if (productIds.length === 0) {
      return NextResponse.json({ error: '제품을 선택해주세요' }, { status: 400 });
    }
    if (productIds.length > 500) {
      return NextResponse.json({ error: '한 번에 500개까지 처리할 수 있습니다' }, { status: 400 });
    }

    // 존재하지 않는 로케이션에 배정하려는 실수 차단 (FK 에러를 그대로 노출하지 않고 친절히)
    if (locationId) {
      const { data: loc, error: locErr } = await supabase
        .from('warehouse_locations')
        .select('id')
        .eq('id', locationId)
        .maybeSingle();
      if (locErr) throw locErr;
      if (!loc) return NextResponse.json({ error: '없는 위치입니다 (삭제됐을 수 있습니다)' }, { status: 404 });
    }

    const { data, error } = await supabase
      .from('products')
      .update({ location_id: locationId })   // ← 오직 이 컬럼만
      .in('id', productIds)
      .select('id');

    if (error) throw error;
    return NextResponse.json({ success: true, updated: data?.length || 0 });
  } catch (err) {
    return NextResponse.json({ error: err instanceof Error ? err.message : String(err) }, { status: 500 });
  }
}
