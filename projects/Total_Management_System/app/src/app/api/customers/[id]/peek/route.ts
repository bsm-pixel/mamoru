import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { CONSULTATION_TYPE_LABEL } from '@/lib/utils/format';

/**
 * GET /api/customers/[id]/peek — 고객 미리 보기 카드용 가벼운 요약 (2026-11-21)
 *
 * 목록에서 이름에 마우스를 올릴 때마다 부른다. 고객 상세(api/customers/[id])는 조회 7개 + 전체 판매 합계라 무거워
 * 여기서는 카드에 보이는 것만 — 유형 · 미수금 · 최근 거래 3건 · 마지막 메모 1줄.
 */
export interface CustomerPeekData {
  customer: {
    id: string; name: string; activity_name: string | null; position: string | null;
    customer_type: string; company_name: string | null; outstanding_balance: number;
  };
  recent: Array<{ kind: '판매' | '주문' | '복원수리' | '상담'; label: string; date: string; amount: number | null; cancelled: boolean }>;
  note: { body: string; created_at: string } | null;
}

export async function GET(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const [cRes, sRes, oRes, rRes, kRes, nRes] = await Promise.all([
      db.from('customers').select('id, name, activity_name, position, customer_type, company_name, outstanding_balance').eq('id', id).single(),
      db.from('offline_sales').select('sale_number, sale_date, paid_amount, cancelled_at, returned_at').eq('customer_id', id).order('sale_date', { ascending: false }).limit(3),
      db.from('orders').select('imweb_order_no, ordered_at, paid_amount, status').eq('customer_id', id).order('ordered_at', { ascending: false }).limit(3),
      db.from('repairs').select('as_id, status, total_amount, created_at').eq('customer_id', id).order('created_at', { ascending: false }).limit(3),
      db.from('consultations').select('consultation_type, visit_date, status, created_at').eq('customer_id', id).order('created_at', { ascending: false }).limit(3),
      db.from('customer_notes').select('body, created_at').eq('customer_id', id).is('deleted_at', null).order('created_at', { ascending: false }).limit(1),
    ]);
    if (cRes.error || !cRes.data) return NextResponse.json({ error: '고객을 찾을 수 없습니다' }, { status: 404 });

    const recent: CustomerPeekData['recent'] = [
      ...((sRes.data || []) as Array<{ sale_number: string; sale_date: string; paid_amount: number | null; cancelled_at: string | null; returned_at: string | null }>)
        .map((s) => ({ kind: '판매' as const, label: s.sale_number, date: s.sale_date, amount: s.paid_amount ?? null, cancelled: !!(s.cancelled_at || s.returned_at) })),
      ...((oRes.data || []) as Array<{ imweb_order_no: string; ordered_at: string; paid_amount: number | null; status: string }>)
        .map((o) => ({ kind: '주문' as const, label: o.imweb_order_no, date: o.ordered_at, amount: o.paid_amount ?? null, cancelled: /cancel/.test(o.status || '') })),
      ...((rRes.data || []) as Array<{ as_id: string; status: string; total_amount: number | null; created_at: string }>)
        .map((r) => ({ kind: '복원수리' as const, label: r.as_id, date: r.created_at, amount: r.total_amount ?? null, cancelled: r.status === 'cancelled' })),
      ...((kRes.data || []) as Array<{ consultation_type: string; visit_date: string | null; status: string; created_at: string }>)
        .map((k) => ({ kind: '상담' as const, label: CONSULTATION_TYPE_LABEL[k.consultation_type] || k.consultation_type, date: k.visit_date || k.created_at, amount: null, cancelled: k.status === 'cancelled' })),
    ].filter((t) => !!t.date).sort((a, b) => String(b.date).localeCompare(String(a.date))).slice(0, 3);

    const body: CustomerPeekData = {
      customer: { ...cRes.data, outstanding_balance: Number(cRes.data.outstanding_balance) || 0 },
      recent,
      note: (nRes.data || [])[0] || null,
    };
    return NextResponse.json(body);
  } catch (err) {
    return NextResponse.json({ error: errMsg(err) }, { status: 500 });
  }
}
