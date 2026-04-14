import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/deliveries — 납품 목록 */
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
      .from('deliveries')
      .select('*', { count: 'exact' })
      .order('created_at', { ascending: false })
      .range(from, to);

    if (status) {
      if (status === 'cancelled') {
        query = query.not('cancelled_at', 'is', null);
      } else {
        query = query.eq('status', status).is('cancelled_at', null);
      }
    } else {
      // 전체: 취소 건 제외
      query = query.is('cancelled_at', null);
    }

    if (search) {
      query = query.or(`customer_name.ilike.%${search}%,dl_number.ilike.%${search}%`);
    }

    const dateRange = url.searchParams.get('date_range');
    if (dateRange && dateRange !== 'all') {
      const now = new Date();
      let dateFrom: string;
      if (dateRange === 'today') dateFrom = now.toISOString().slice(0, 10);
      else if (dateRange === 'week') { const d = new Date(now); d.setDate(d.getDate() - 7); dateFrom = d.toISOString().slice(0, 10); }
      else { const d = new Date(now); d.setMonth(d.getMonth() - 1); dateFrom = d.toISOString().slice(0, 10); }
      query = query.gte('delivery_date', dateFrom);
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ deliveries: data || [], total: count || 0 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/deliveries — 납품서 생성 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const {
      customer_id, customer_name, customer_phone, customer_type,
      delivery_date, expected_date, memo, vat_type, receipt_type,
      payment_status, payment_method, paid_amount, discount_amount,
      items,
    } = body as {
      customer_id?: string;
      customer_name: string;
      customer_phone?: string;
      customer_type?: string;
      delivery_date?: string;
      expected_date?: string;
      memo?: string;
      vat_type?: 'included' | 'separate' | 'none';
      receipt_type?: string;
      payment_status?: string;
      payment_method?: string;
      paid_amount?: number;
      discount_amount?: number;
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        category?: string;
        quantity: number;
        unit_price: number;
      }>;
    };

    if (!customer_name?.trim()) {
      return NextResponse.json({ error: '고객명은 필수입니다' }, { status: 400 });
    }
    if (!items || items.length === 0) {
      return NextResponse.json({ error: '품목을 추가해주세요' }, { status: 400 });
    }

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // 납품번호 생성: DL-YYYYMMDD-NNN
    const dlDate = delivery_date || new Date().toISOString().slice(0, 10);
    const dateStr = dlDate.replace(/-/g, '');
    const { count } = await db
      .from('deliveries')
      .select('*', { count: 'exact', head: true })
      .gte('delivery_date', dlDate);
    const seq = String((count || 0) + 1).padStart(3, '0');
    const dlNumber = `DL-${dateStr}-${seq}`;

    // 합계 계산
    const vatTypeVal = vat_type || 'included';
    const itemTotal = items.reduce((s, i) => s + i.quantity * i.unit_price, 0);
    const discountVal = discount_amount || 0;
    const baseAmount = itemTotal - discountVal;
    let supplyAmount = 0, vatAmount = 0, totalAmount = 0;

    if (vatTypeVal === 'separate') {
      supplyAmount = baseAmount;
      vatAmount = Math.round(baseAmount * 0.1);
      totalAmount = baseAmount + vatAmount;
    } else if (vatTypeVal === 'none') {
      supplyAmount = baseAmount;
      vatAmount = 0;
      totalAmount = baseAmount;
    } else {
      supplyAmount = Math.round(baseAmount / 1.1);
      vatAmount = baseAmount - supplyAmount;
      totalAmount = baseAmount;
    }

    const paidVal = paid_amount || 0;
    const payStatus = payment_status || 'unpaid';

    // 복원수리(RS) 전용 납품은 재고 차감 없으므로 즉시 confirmed
    const isRepairOnly = items.every((i) => i.category === 'RS');

    const { data: delivery, error: dlError } = await db
      .from('deliveries')
      .insert({
        dl_number: dlNumber,
        customer_id: customer_id || null,
        customer_name: customer_name.trim(),
        customer_phone: customer_phone || null,
        customer_type: customer_type || null,
        delivery_date: dlDate,
        expected_date: expected_date || null,
        total_amount: totalAmount,
        discount_amount: discountVal,
        paid_amount: paidVal,
        payment_status: payStatus,
        payment_method: payment_method || 'transfer',
        vat_type: vatTypeVal,
        supply_amount: supplyAmount,
        vat_amount: vatAmount,
        receipt_type: receipt_type || 'expense_proof',
        memo: memo || null,
        created_by: user.id,
        ...(isRepairOnly ? { status: 'confirmed' } : {}),
      })
      .select()
      .single();

    if (dlError) throw dlError;

    // 품목 생성
    const dlItems = items.map((item) => ({
      delivery_id: delivery.id,
      product_id: item.product_id || null,
      product_name: item.product_name,
      sku: item.sku || null,
      category: item.category || null,
      quantity: item.quantity,
      unit_price: item.unit_price,
      total_price: item.quantity * item.unit_price,
    }));

    const { error: itemsError } = await db.from('delivery_items').insert(dlItems);
    if (itemsError) throw itemsError;

    // 미수금 반영 (미결제/부분결제 시)
    if (customer_id && payStatus !== 'paid') {
      const unpaid = totalAmount - paidVal;
      if (unpaid > 0) {
        const { data: cust } = await db.from('customers').select('outstanding_balance').eq('id', customer_id).single();
        if (cust) {
          await db.from('customers').update({
            outstanding_balance: (cust.outstanding_balance || 0) + unpaid,
          }).eq('id', customer_id);
        }
      }
    }

    return NextResponse.json({ delivery, dlNumber });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
