import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

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

    const [customerRes, salesRes, contractsRes, consultationsRes, repairsRes] = await Promise.all([
      db.from('customers').select('*').eq('id', id).single(),
      db.from('offline_sales')
        .select('id, sale_number, sale_date, total_amount, paid_amount, payment_method, payment_status')
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
    ]);

    if (customerRes.error) throw customerRes.error;

    // 판매 합계 계산
    const sales = (salesRes.data || []) as Array<{ paid_amount: number; sale_date: string; cancelled_at?: string }>;
    const activeSales = sales.filter((s) => !s.cancelled_at);
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
      summary: {
        totalSales: activeSales.length,
        totalSalesAmount,
        lastSaleDate,
      },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
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
    const allowed = ['name', 'phone', 'email', 'postcode', 'address_road', 'address_detail', 'customer_type', 'company_name', 'memo', 'outstanding_balance'];
    for (const key of allowed) {
      if (key in body) updates[key] = body[key];
    }

    if (Object.keys(updates).length === 0) {
      return NextResponse.json({ error: '수정할 항목이 없습니다' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const { data: customer, error } = await db
      .from('customers')
      .update(updates)
      .eq('id', id)
      .select()
      .single();

    if (error) throw error;

    return NextResponse.json({ customer });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
