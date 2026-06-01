import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/sourcing/by-product?product_id=xxx
 * 제품/부자재 상세에서 연결된 소싱 정보(업체명·링크·단가) 역조회.
 * linked_product_id = product_id 인 소싱 품목 최신 1건 + 회차 환율로 한화 환산.
 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const productId = req.nextUrl.searchParams.get('product_id');
    if (!productId) return NextResponse.json({ sourcing: null });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { data: items, error } = await db
      .from('sourcing_items')
      .select('id, sticker_no, supplier_name, supplier_url, vendor_url, unit_price, features_memo, po_id, selected_at')
      .eq('linked_product_id', productId)
      .order('selected_at', { ascending: false, nullsFirst: false })
      .limit(1);
    if (error) throw error;

    const item = items?.[0];
    if (!item) return NextResponse.json({ sourcing: null });

    // 회차 환율 → 한화 환산
    let exchangeRate = 0;
    let poNumber: string | null = null;
    if (item.po_id) {
      const { data: po } = await db
        .from('sourcing_pos')
        .select('po_number, exchange_rate')
        .eq('id', item.po_id)
        .single();
      exchangeRate = po?.exchange_rate || 0;
      poNumber = po?.po_number ?? null;
    }

    return NextResponse.json({
      sourcing: {
        sticker_no: item.sticker_no,
        supplier_name: item.supplier_name,
        supplier_url: item.supplier_url,
        vendor_url: item.vendor_url,
        unit_price: item.unit_price,
        features_memo: item.features_memo,
        krw_price: Math.round((item.unit_price || 0) * exchangeRate),
        exchange_rate: exchangeRate,
        po_number: poNumber,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
