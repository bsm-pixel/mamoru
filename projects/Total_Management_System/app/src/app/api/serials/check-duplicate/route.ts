import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/serials/check-duplicate?serial={번호}&excludeSaleId={현재 판매 id}
 *
 * 시리얼 번호로 기존 등록 여부 확인. 다른 판매에 이미 등록되어 있으면
 * 그 판매의 정보(sale_number·customer·product·sale_date)를 반환.
 *
 * 사용처: SerialPicker — 사장님이 시리얼 직접 입력·자동 생성 시 즉시 호출.
 *         반환에 exists=true 면 다이얼로그로 사장님 명시 동의 받은 후 진행.
 *
 * 2026-05-18 Phase A 신규 — 시리얼 사일런트 강탈 방지.
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const { searchParams } = new URL(req.url);
    const serialNumber = searchParams.get('serial')?.trim();
    const excludeSaleId = searchParams.get('excludeSaleId') || null;

    if (!serialNumber) {
      return NextResponse.json({ error: 'serial 파라미터 필요' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 시리얼 번호 일치하는 모든 시리얼 조회 (sold + in_stock 모두 포함)
    const { data: serials } = await db
      .from('product_serials')
      .select('id, serial_number, status, offline_sale_id, sale_item_id, sold_to_name, sold_at, product_id')
      .eq('serial_number', serialNumber)
      .limit(5);

    if (!serials || serials.length === 0) {
      return NextResponse.json({ exists: false });
    }

    // 현재 판매 (excludeSaleId) 이외의 다른 판매에 등록된 케이스만 충돌로 본다
    const conflict = serials.find((s: { offline_sale_id: string | null }) =>
      s.offline_sale_id && s.offline_sale_id !== excludeSaleId
    );

    if (!conflict) {
      // 재고에 있는 시리얼이거나 같은 판매에 있는 케이스 — 충돌 아님
      return NextResponse.json({ exists: false });
    }

    // 충돌 상세 정보 조회 (sale_number + product_name)
    const { data: sale } = await db
      .from('offline_sales')
      .select('sale_number, sale_date, customer_id')
      .eq('id', conflict.offline_sale_id)
      .single();

    let productName: string | null = null;
    if (conflict.sale_item_id) {
      const { data: item } = await db
        .from('offline_sale_items')
        .select('product_name')
        .eq('id', conflict.sale_item_id)
        .single();
      productName = item?.product_name || null;
    }

    return NextResponse.json({
      exists: true,
      serial_id: conflict.id,
      serial_number: conflict.serial_number,
      status: conflict.status,
      sale_number: sale?.sale_number || null,
      sale_date: sale?.sale_date || null,
      customer_name: conflict.sold_to_name || null,
      product_name: productName,
    });
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 500 });
  }
}
