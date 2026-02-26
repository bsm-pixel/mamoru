import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/contracts — 계약서 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const url = req.nextUrl;
    const status = url.searchParams.get('status');
    const search = url.searchParams.get('search');
    const page = parseInt(url.searchParams.get('page') || '1');
    const limit = parseInt(url.searchParams.get('limit') || '20');
    const from = (page - 1) * limit;
    const to = from + limit - 1;

    let query = db
      .from('contracts')
      .select('*', { count: 'exact' })
      .order('created_at', { ascending: false })
      .range(from, to);

    if (status && status !== 'all') {
      query = query.eq('status', status);
    }
    if (search) {
      query = query.or(
        `customer_name.ilike.%${search}%,customer_phone.ilike.%${search}%,contract_number.ilike.%${search}%`
      );
    }

    const { data, count, error } = await query;
    if (error) throw error;

    return NextResponse.json({ contracts: data || [], total: count || 0 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/contracts — 계약서 생성 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const body = await req.json();
    const { contract, items } = body;

    // 계약번호: CT-YYYYMMDD-NNN
    const today = new Date().toISOString().slice(0, 10);
    const todayCompact = today.replace(/-/g, '');
    const { count } = await db
      .from('contracts')
      .select('*', { count: 'exact', head: true })
      .gte('created_at', `${today}T00:00:00`);
    const seq = String((count || 0) + 1).padStart(3, '0');
    const contractNumber = `CT-${todayCompact}-${seq}`;

    // 서명 여부로 상태 결정
    const status = contract.signature_data ? 'signed' : 'draft';

    const { data: created, error: createError } = await db
      .from('contracts')
      .insert({
        contract_number: contractNumber,
        customer_name: contract.customer_name,
        customer_phone: contract.customer_phone || null,
        customer_email: contract.customer_email || null,
        customer_address: contract.customer_address || null,
        total_amount: contract.total_amount,
        discount_amount: contract.discount_amount || 0,
        final_amount: contract.final_amount,
        payment_method: contract.payment_method,
        installment_months: contract.installment_months || 0,
        signature_data: contract.signature_data || null,
        signed_at: contract.signature_data ? new Date().toISOString() : null,
        status,
        memo: contract.memo || null,
        created_by: user.id,
      })
      .select()
      .single();

    if (createError) throw createError;

    // 항목 생성
    if (items && items.length > 0) {
      const contractItems = items.map((item: Record<string, unknown>) => ({
        contract_id: created.id,
        product_id: item.product_id || null,
        product_name: item.product_name,
        sku: item.sku || null,
        quantity: item.quantity,
        unit_price: item.unit_price,
        total_price: item.total_price,
        option_text: item.option_text || null,
      }));

      const { error: itemsError } = await db
        .from('contract_items')
        .insert(contractItems);

      if (itemsError) throw itemsError;
    }

    return NextResponse.json({ contract: created, contractNumber });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
