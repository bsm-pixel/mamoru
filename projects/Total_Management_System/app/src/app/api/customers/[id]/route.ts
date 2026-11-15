import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { phoneError, emailError } from '@/lib/utils/validate';

/** GET /api/customers/[id] — 고객 상세 (판매내역 포함) */
export async function GET(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const [customerRes, salesRes, contractsRes, consultationsRes, repairsRes, ordersRes, salesAllRes] = await Promise.all([
      db.from('customers').select('*').eq('id', id).single(),
      db.from('offline_sales')
        .select('id, sale_number, sale_date, total_amount, paid_amount, payment_method, payment_status, cancelled_at, returned_at')
        .eq('customer_id', id)
        .order('sale_date', { ascending: false })
        .limit(20),
      db.from('contracts')
        .select('id, contract_number, final_amount, status, created_at')
        .eq('customer_id', id)
        .order('created_at', { ascending: false })
        .limit(10),
      db.from('consultations')
        .select('id, consultation_type, visit_date, status, created_at')
        .eq('customer_id', id)
        .order('created_at', { ascending: false })
        .limit(10),
      db.from('repairs')
        .select('id, repair_number, status, item_description, total_cost, created_at')
        .eq('customer_id', id)
        .order('created_at', { ascending: false })
        .limit(10),
      // 128: 아임웹 주문도 고객 타임라인에 포함 (customer_id 연결 후)
      db.from('orders')
        .select('id, imweb_order_no, ordered_at, paid_amount, status')
        .eq('customer_id', id)
        .order('ordered_at', { ascending: false })
        .limit(20),
      // 2026-11-15 합계 전용 — 위 목록은 최근 20건만이라 합계에 쓰면 21건째부터 빠진다. 취소·반품 판정 칸 포함
      db.from('offline_sales')
        .select('paid_amount, sale_date, cancelled_at, returned_at')
        .eq('customer_id', id),
    ]);

    if (customerRes.error) throw customerRes.error;

    // 판매 합계 계산
    //   2026-11-15 취소한 판매가 「판매 N건 · 매출」에 그대로 세던 문제 — 조회에 cancelled_at 이 빠져 걸러지지 않았다.
    //   취소·반품 건 제외(대시보드 use-dashboard-stats 와 같은 기준). 합계는 전체 판매(salesAllRes)로 계산
    const sales = (salesRes.data || []) as Array<{ paid_amount: number; sale_date: string; cancelled_at?: string | null; returned_at?: string | null }>;
    const allSales = (salesAllRes.data || []) as Array<{ paid_amount: number; sale_date: string; cancelled_at?: string | null; returned_at?: string | null }>;
    const activeSales = allSales.filter((s) => !s.cancelled_at && !s.returned_at);
    const totalSalesAmount = activeSales.reduce((s, r) => s + (r.paid_amount || 0), 0);
    const lastSaleDate = activeSales.length > 0
      ? activeSales.sort((a, b) => b.sale_date.localeCompare(a.sale_date))[0].sale_date
      : null;

    return NextResponse.json({
      customer: customerRes.data,
      sales,
      contracts: contractsRes.data || [],
      consultations: consultationsRes.data || [],
      repairs: repairsRes.data || [],
      orders: ordersRes.data || [],
      summary: {
        totalSales: activeSales.length,
        totalSalesAmount,
        lastSaleDate,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}

/** PATCH /api/customers/[id] — 고객 정보 수정 */
export async function PATCH(
  req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const updates: Record<string, unknown> = {};

    // 허용 필드만 추출
    const allowed = ['name', 'phone', 'email', 'postcode', 'address_road', 'address_detail', 'customer_type', 'company_name', 'activity_name', 'position', 'memo', 'outstanding_balance', 'tags', 'default_repair_price', 'business_number', 'representative', 'business_type', 'business_category', 'contact_channel'];
    for (const key of allowed) {
      if (key in body) updates[key] = body[key];
    }

    if (Object.keys(updates).length === 0) {
      return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
    }
    // 2026-11-15 서버에서도 검사 (바꾸려는 칸만)
    const inputErr = ('phone' in updates ? phoneError(updates.phone) : null) || ('email' in updates ? emailError(updates.email) : null)
      || ('name' in updates && !String(updates.name ?? '').trim() ? '고객명은 비워 둘 수 없습니다' : null);
    if (inputErr) return NextResponse.json({ error: inputErr }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: customer, error } = await db
      .from('customers')
      .update(updates)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    // 고객명/연락처 변경 시 → 해당 고객의 판매 건도 일괄 업데이트 (송장 등에서 최신 정보 사용)
    if (updates.name || updates.phone) {
      const saleUpdates: Record<string, unknown> = {};
      if (updates.name) saleUpdates.customer_name = updates.name;
      if (updates.phone) saleUpdates.customer_phone = updates.phone;
      await db.from('offline_sales').update(saleUpdates).eq('customer_id', id);
    }

    return NextResponse.json({ customer });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}

/**
 * DELETE /api/customers/[id] — 고객 삭제 (2026-11-15)
 * 거래·이력이 하나도 없는 고객만 지운다. 하나라도 있으면 무엇이 걸려 있는지 사람 말로 알려 주고 막는다.
 * (잘못 만든 고객·중복 등록을 정리하는 용도. 거래가 있는 고객을 감추는 「숨기기」는 표에 칸이 필요해 따로 진행)
 */
export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: target } = await db.from('customers').select('id, name').eq('id', id).maybeSingle();
    if (!target) return NextResponse.json({ error: '고객을 찾을 수 없습니다' }, { status: 404 });

    // 이 고객을 가리키는 기록 — [표, 칸, 화면에 보일 이름]
    const LINKS: Array<[string, string, string]> = [
      ['offline_sales', 'customer_id', '판매'],
      ['deliveries', 'customer_id', '납품'],
      ['orders', 'customer_id', '주문'],
      ['repairs', 'customer_id', '복원수리'],
      ['consultations', 'customer_id', '상담'],
      ['contracts', 'customer_id', '계약서'],
      ['returns', 'customer_id', '반품·교환'],
      ['event_submissions', 'customer_id', '이벤트 접수'],
      ['purchase_orders', 'supplier_id', '발주'],
      ['products', 'supplier_id', '공급 제품'],
      ['customer_product_catalog', 'customer_id', '거래처 단가'],
      ['supplier_product_catalog', 'supplier_id', '매입 단가'],
      ['customers', 'merged_into_id', '병합된 고객'],
    ];
    const counts = await Promise.all(
      LINKS.map(([t, col]) => db.from(t).select('id', { count: 'exact', head: true }).eq(col, id)),
    );
    const linked: string[] = [];
    counts.forEach((r: { count: number | null; error: unknown }, i: number) => {
      // 조회 자체가 실패하면(칸 이름 차이 등) 안전하게 "있을 수 있다"로 보고 막는다
      if (r.error) linked.push(`${LINKS[i][2]} 확인 실패`);
      else if ((r.count || 0) > 0) linked.push(`${LINKS[i][2]} ${r.count}건`);
    });
    if (linked.length > 0) {
      return NextResponse.json({
        error: `거래·이력이 있는 고객은 삭제할 수 없습니다 (${linked.join(', ')}). 대신 [고객 숨기기]를 쓰면 목록 · 검색에서 빠집니다.`,
        linked,
      }, { status: 400 });
    }

    // 메모는 거래가 아니므로 같이 지운다
    await db.from('customer_notes').delete().eq('customer_id', id);
    const { error } = await db.from('customers').delete().eq('id', id);
    if (error) {
      if (error.code === '23503') {
        return NextResponse.json({ error: '다른 기록이 이 고객을 쓰고 있어 삭제할 수 없습니다.' }, { status: 400 });
      }
      throw error;
    }
    return NextResponse.json({ deleted: true, name: target.name });
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
