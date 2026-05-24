import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/repair/[id]/related-shipments
 *
 * 복원수리 합포장 출고 시 같은 고객의 송장 보유 판매건/주문건을 검색.
 *
 * 흐름:
 *   1. 해당 repair 의 phone 으로 정규화 키 추출
 *   2. offline_sales: 송장 있고 취소 안 된 건 (sale_date DESC)
 *   3. orders (아임웹): 송장 있고 배송 상태 건 (shipped_at DESC)
 *   4. 응답: { sales: [...], orders: [...] }
 *
 * 활용: 합포장 출고 모달에서 사장님이 1초에 매칭할 판매/주문 선택
 */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) {
      return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 1. repair 의 phone 조회
    const { data: repair, error: repairErr } = await db
      .from('repairs')
      .select('id, phone, name')
      .eq('id', id)
      .single();

    if (repairErr || !repair) {
      return NextResponse.json({ error: '복원수리 건을 찾을 수 없습니다' }, { status: 404 });
    }

    // 전화번호 정규화 (숫자만 추출)
    const phoneNormalized = (repair.phone || '').replace(/\D/g, '');
    if (!phoneNormalized) {
      return NextResponse.json({ sales: [], orders: [], phone: repair.phone, name: repair.name });
    }

    // 2. offline_sales 검색 (송장 있는 건만, 취소 제외)
    //    phone 정규화 비교를 위해 customer_phone 자체에 normalize 적용
    const [salesRes, ordersRes] = await Promise.all([
      db.from('offline_sales')
        .select('id, sale_number, customer_name, customer_phone, sale_date, total_amount, invoice_number, courier_name, shipped_at')
        .not('invoice_number', 'is', null)
        .is('cancelled_at', null)
        .order('sale_date', { ascending: false })
        .limit(50),
      db.from('orders')
        .select('id, imweb_order_no, orderer_name, orderer_phone, total_price, invoice_number, shipped_at, status, paid_at')
        .not('invoice_number', 'is', null)
        .in('status', ['shipping', 'delivered'])
        .order('shipped_at', { ascending: false })
        .limit(50),
    ]);

    // phone 정규화 후 필터 (DB 단에 generated 컬럼 없을 수 있어 JS 필터)
    const matchPhone = (raw: string | null) =>
      (raw || '').replace(/\D/g, '') === phoneNormalized;

    const sales = (salesRes.data || []).filter((s: { customer_phone: string | null }) => matchPhone(s.customer_phone));
    const orders = (ordersRes.data || []).filter((o: { orderer_phone: string | null }) => matchPhone(o.orderer_phone));

    return NextResponse.json({
      phone: repair.phone,
      name: repair.name,
      sales,
      orders,
    });
  } catch (err) {
    console.error('[related-shipments] 조회 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
