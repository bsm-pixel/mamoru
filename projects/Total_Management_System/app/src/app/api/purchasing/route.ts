import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/purchasing — 발주 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const url = req.nextUrl;
    const status = url.searchParams.get('status');
    const search = url.searchParams.get('search');
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    let query = db
      .from('purchase_orders')
      .select('*', { count: 'exact' })
      .order('created_at', { ascending: false })
      .range(from, to);

    if (status) query = query.eq('status', status);
    if (search) {
      query = query.or(`supplier_name.ilike.%${search}%,po_number.ilike.%${search}%`);
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ orders: data || [], total: count || 0 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/purchasing — 발주 생성 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { supplier_id, supplier_name, order_date, expected_date, memo, items } = body as {
      supplier_id?: string;
      supplier_name: string;
      order_date?: string;
      expected_date?: string;
      memo?: string;
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
      }>;
    };

    if (!supplier_name?.trim()) {
      return NextResponse.json({ error: '매입처명은 필수입니다' }, { status: 400 });
    }
    if (!items || items.length === 0) {
      return NextResponse.json({ error: '품목을 추가해주세요' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 발주번호 생성: PO-YYYYMMDD-NNN
    const poDate = order_date || new Date().toISOString().slice(0, 10);
    const dateStr = poDate.replace(/-/g, '');
    const { count } = await db
      .from('purchase_orders')
      .select('*', { count: 'exact', head: true })
      .gte('order_date', poDate);
    const seq = String((count || 0) + 1).padStart(3, '0');
    const poNumber = `PO-${dateStr}-${seq}`;

    // 합계 계산
    const totalAmount = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
    const supplyAmount = Math.round(totalAmount / 1.1);
    const vatAmount = totalAmount - supplyAmount;

    const { data: order, error: poError } = await db
      .from('purchase_orders')
      .insert({
        po_number: poNumber,
        supplier_id: supplier_id || null,
        supplier_name: supplier_name.trim(),
        order_date: poDate,
        expected_date: expected_date || null,
        total_amount: totalAmount,
        balance_amount: totalAmount,
        is_vat_included: true,
        supply_amount: supplyAmount,
        vat_amount: vatAmount,
        memo: memo || null,
        created_by: user.id,
      })
      .select()
      .single();

    if (poError) throw poError;

    // 품목 생성
    const poItems = items.map((item) => ({
      po_id: order.id,
      product_id: item.product_id || null,
      product_name: item.product_name,
      sku: item.sku || null,
      quantity: item.quantity,
      unit_price: item.unit_price,
      total_price: item.quantity * item.unit_price,
    }));

    const { error: itemsError } = await db
      .from('purchase_order_items')
      .insert(poItems);

    if (itemsError) throw itemsError;

    return NextResponse.json({ order, poNumber });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
