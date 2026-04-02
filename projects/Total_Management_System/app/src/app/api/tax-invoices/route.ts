import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** GET /api/tax-invoices — 세금계산서 목록 */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const sp = req.nextUrl.searchParams;
    const from = sp.get('from');
    const to = sp.get('to');
    const type = sp.get('type'); // sales | purchase

    let query = db.from('tax_invoices').select('*').order('issue_date', { ascending: false }).limit(200);
    if (from) query = query.gte('issue_date', from);
    if (to) query = query.lte('issue_date', to);
    if (type) query = query.eq('invoice_type', type);

    const { data, error } = await query;
    if (error) throw error;

    const items = data || [];
    const salesTotal = items.filter((i: { invoice_type: string }) => i.invoice_type === 'sales').reduce((s: number, i: { total_amount: number }) => s + i.total_amount, 0);
    const purchaseTotal = items.filter((i: { invoice_type: string }) => i.invoice_type === 'purchase').reduce((s: number, i: { total_amount: number }) => s + i.total_amount, 0);
    const salesTax = items.filter((i: { invoice_type: string }) => i.invoice_type === 'sales').reduce((s: number, i: { tax_amount: number }) => s + i.tax_amount, 0);
    const purchaseTax = items.filter((i: { invoice_type: string }) => i.invoice_type === 'purchase').reduce((s: number, i: { tax_amount: number }) => s + i.tax_amount, 0);

    return NextResponse.json({
      invoices: items,
      summary: { salesTotal, purchaseTotal, salesTax, purchaseTax, netTax: salesTax - purchaseTax, count: items.length },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/tax-invoices — 세금계산서 등록 */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const { invoice_type, issue_date, counterparty_name, counterparty_biz_no, supply_amount, tax_amount, sale_id, purchase_order_id, memo } = body;

    if (!counterparty_name) return NextResponse.json({ error: '거래처명 필수' }, { status: 400 });

    const total = (supply_amount || 0) + (tax_amount || 0);

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (supabase as any).from('tax_invoices').insert({
      invoice_type: invoice_type || 'sales',
      issue_date: issue_date || new Date().toISOString().slice(0, 10),
      counterparty_name,
      counterparty_biz_no: counterparty_biz_no || null,
      supply_amount: supply_amount || 0,
      tax_amount: tax_amount || 0,
      total_amount: total,
      sale_id: sale_id || null,
      purchase_order_id: purchase_order_id || null,
      memo: memo || null,
      created_by: user.id,
    }).select().single();

    if (error) throw error;
    return NextResponse.json(data);
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE /api/tax-invoices?id=xxx */
export async function DELETE(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const id = req.nextUrl.searchParams.get('id');
    if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await (supabase as any).from('tax_invoices').delete().eq('id', id);
    return NextResponse.json({ ok: true });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
