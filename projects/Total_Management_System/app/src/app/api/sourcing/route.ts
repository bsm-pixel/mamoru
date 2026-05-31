import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/** KST 기준 YYYYMMDD (Vercel UTC 보정) */
function kstYmd(): string {
  return new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10).replaceAll('-', '');
}

/** GET /api/sourcing — 소싱 발주 목록 (품목 수 포함) */
export async function GET(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const status = req.nextUrl.searchParams.get('status');
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    let query = db
      .from('sourcing_pos')
      .select('*, sourcing_items(id, inspection_status)')
      .order('created_at', { ascending: false });
    if (status) query = query.eq('status', status);

    const { data, error } = await query;
    if (error) throw error;

    // 품목 상태 집계 부착
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const orders = (data || []).map((po: any) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items: any[] = po.sourcing_items || [];
      const counts = { total: items.length, pending: 0, matched: 0, selected: 0, rejected: 0 };
      for (const it of items) {
        if (it.inspection_status in counts) counts[it.inspection_status as keyof typeof counts]++;
      }
      const { sourcing_items: _drop, ...rest } = po;
      void _drop;
      return { ...rest, counts };
    });

    return NextResponse.json({ orders });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** POST /api/sourcing — 소싱 발주 생성 (헤더 + 품목 일괄) */
export async function POST(req: NextRequest) {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    // po_number 자동 채번: SRC-YYYYMMDD-NNN
    const ymd = kstYmd();
    const prefix = `SRC-${ymd}`;
    const { data: todays } = await db
      .from('sourcing_pos')
      .select('po_number')
      .like('po_number', `${prefix}-%`);
    const seq = String((todays?.length || 0) + 1).padStart(3, '0');
    const po_number = `${prefix}-${seq}`;

    const { data: po, error: poErr } = await db
      .from('sourcing_pos')
      .insert({
        po_number,
        supplier_name: body.supplier_name || null,
        supplier_url: body.supplier_url || null,
        order_date: body.order_date || ymd.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3'),
        exchange_rate: body.exchange_rate ?? 195,
        memo: body.memo || null,
        created_by: user.id,
      })
      .select()
      .single();
    if (poErr) throw poErr;

    // 품목 일괄 INSERT (sticker_no 자동: {po_number}-{seq})
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const rawItems: any[] = Array.isArray(body.items) ? body.items : [];
    if (rawItems.length > 0) {
      const itemsToInsert = rawItems.map((it, idx) => ({
        po_id: po.id,
        sticker_no: `${po_number}-${String(idx + 1).padStart(3, '0')}`,
        supplier_name: it.supplier_name || null,
        supplier_url: it.supplier_url || null,
        vendor_url: it.vendor_url || null,
        product_name: it.product_name || '',
        features_memo: it.features_memo || null,
        unit_price: it.unit_price ?? 0,
        moq: it.moq ?? null,
        sort_order: idx,
      }));
      const { error: itemErr } = await db.from('sourcing_items').insert(itemsToInsert);
      if (itemErr) throw itemErr;
    }

    return NextResponse.json({ id: po.id, po_number }, { status: 201 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
