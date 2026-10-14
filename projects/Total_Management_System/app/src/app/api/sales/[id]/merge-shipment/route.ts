import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { COURIER_LOTTE, isAlpsTrackable } from '@/lib/shipping/couriers';

/**
 * 판매 합포장 (162, 2026-10-14) — 납품 합포장(158)과 같은 규칙
 *
 * GET  /api/sales/[id]/merge-shipment → { mergedInto, mergedChildren, candidates }
 * POST /api/sales/[id]/merge-shipment { target_id } → 이 판매를 target 판매의 송장에 얹는다
 *
 * 얹힌 건은 출고 전 상태(shipped_at NULL) 그대로 둔다. 크론 [3-A]가 판매 건마다 송장번호로 집하를 조회하므로
 * 원 송장이 수거되면 같은 회차에 함께 출고완료 + **각자 전화번호로 출고 알림톡**, 배송완료·후기 요청도 건마다 따로.
 * 그래서 대상은 자동추적되는 롯데 송장만 받는다 (타 택배사면 출고·알림톡이 영영 안 남).
 */

const COLS = 'id, sale_number, customer_id, customer_name, customer_phone, sale_date, invoice_number, courier_name, shipped_at, ship_address_road, ship_address_detail';

type SaleRow = {
  id: string; sale_number: string; customer_id: string | null; customer_name: string; customer_phone: string | null;
  sale_date: string; invoice_number: string | null; courier_name: string | null; shipped_at: string | null;
  ship_address_road?: string | null; ship_address_detail?: string | null;
};

/** 주소 비교용 — 공백 제거 */
const norm = (v?: string | null) => (v || '').replace(/\s+/g, '');

export async function GET(_req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    // 마이그 162 전이면 컬럼이 없어 에러 → 빈 값 (기존 화면 무영향)
    const { data: sale, error } = await db.from('offline_sales')
      .select(`${COLS}, merged_into_sale_id, cancelled_at, returned_at, delivered_at`)
      .eq('id', id).single();
    if (error || !sale) return NextResponse.json({ mergedInto: null, mergedChildren: [], candidates: [] });

    let mergedInto: SaleRow | null = null;
    if (sale.merged_into_sale_id) {
      const { data } = await db.from('offline_sales').select(COLS).eq('id', sale.merged_into_sale_id).maybeSingle();
      mergedInto = data || null;
    }
    const { data: children } = await db.from('offline_sales').select(COLS)
      .eq('merged_into_sale_id', id).is('cancelled_at', null);

    // 후보: 이 판매가 송장 없음·진행 중일 때만 — 최근 30일, 롯데 송장 있고 배송 안 끝난 원 송장 판매
    let candidates: Array<SaleRow & { same_address: boolean; same_customer: boolean; address: string }> = [];
    if (!sale.invoice_number && !sale.cancelled_at && !sale.returned_at && !sale.delivered_at) {
      const since = new Date(Date.now() - 30 * 86400_000).toISOString().slice(0, 10);
      const { data: rows } = await db.from('offline_sales').select(COLS)
        .neq('id', id)
        .not('invoice_number', 'is', null)
        .is('merged_into_sale_id', null)   // 이미 얹혀 가는 건은 원본이 아니다
        .is('cancelled_at', null).is('returned_at', null).is('delivered_at', null)
        .gte('sale_date', since)
        .order('sale_date', { ascending: false })
        .limit(50);
      const list: SaleRow[] = (rows || []).filter((r: SaleRow) => isAlpsTrackable(r.courier_name));

      // 주소 = 별도 배송지(159) 우선, 없으면 고객 등록 주소
      const custIds = [...new Set([sale.customer_id, ...list.map((r) => r.customer_id)].filter(Boolean))];
      const addrMap: Record<string, { road: string | null; detail: string | null }> = {};
      if (custIds.length > 0) {
        const { data: custs } = await db.from('customers').select('id, address_road, address_detail').in('id', custIds);
        for (const c of custs || []) addrMap[c.id] = { road: c.address_road, detail: c.address_detail };
      }
      const addrOf = (r: SaleRow) => r.ship_address_road
        ? { road: r.ship_address_road, detail: r.ship_address_detail || null }
        : (r.customer_id ? addrMap[r.customer_id] : undefined) || { road: null, detail: null };
      const mine = addrOf(sale);

      candidates = list.map((r) => {
        const a = addrOf(r);
        return {
          ...r,
          address: [a.road, a.detail].filter(Boolean).join(' '),
          same_address: !!mine.road && norm(mine.road) === norm(a.road) && norm(mine.detail) === norm(a.detail),
          same_customer: !!sale.customer_id && sale.customer_id === r.customer_id,
        };
      })
        // 같은 주소 → 같은 고객 → 최신순
        .sort((a, b) => Number(b.same_address) - Number(a.same_address) || Number(b.same_customer) - Number(a.same_customer))
        .slice(0, 15);
    }

    return NextResponse.json({ mergedInto, mergedChildren: children || [], candidates });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

export async function POST(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;
    const { target_id } = await req.json() as { target_id?: string };
    if (!target_id || target_id === id) return NextResponse.json({ error: '같이 보낼 판매 건을 골라 주세요' }, { status: 400 });

    const { data: sale } = await db.from('offline_sales').select('*').eq('id', id).single();
    if (!sale) return NextResponse.json({ error: '판매 건을 찾을 수 없습니다' }, { status: 404 });
    if (sale.cancelled_at || sale.returned_at) return NextResponse.json({ error: '취소·반품된 판매입니다' }, { status: 400 });
    if (sale.delivered_at) return NextResponse.json({ error: '이미 수령·배송완료된 판매입니다' }, { status: 400 });
    if (sale.invoice_number) return NextResponse.json({ error: '이미 송장이 있습니다. 먼저 송장을 취소해 주세요' }, { status: 400 });

    const { data: target } = await db.from('offline_sales').select('*').eq('id', target_id).single();
    if (!target) return NextResponse.json({ error: '같이 보낼 판매 건을 찾을 수 없습니다' }, { status: 404 });
    if (!target.invoice_number || target.cancelled_at || target.returned_at || target.delivered_at) {
      return NextResponse.json({ error: '그 판매의 송장이 없거나 이미 배송이 끝났습니다' }, { status: 400 });
    }
    if (target.merged_into_sale_id) {
      return NextResponse.json({ error: '그 판매도 다른 송장에 얹혀 가는 건입니다. 원래 송장 건을 골라 주세요' }, { status: 400 });
    }
    if (!isAlpsTrackable(target.courier_name)) {
      return NextResponse.json({ error: '롯데 송장에만 같이 보낼 수 있습니다 (다른 택배사는 자동 출고·알림톡이 안 됩니다)' }, { status: 400 });
    }

    const { error: mErr } = await db.from('offline_sales').update({
      invoice_number: target.invoice_number,
      courier_name: target.courier_name || COURIER_LOTTE,
      delivery_method: 'shipping',
      merged_into_sale_id: target.id,
    }).eq('id', id).is('invoice_number', null);
    if (mErr) {
      return NextResponse.json({ error: `합포장 저장 실패 — 마이그레이션 162를 실행했는지 확인해 주세요 (${mErr.message})` }, { status: 500 });
    }

    return NextResponse.json({ success: true, mergedInto: target.sale_number, targetName: target.customer_name, invoiceNumber: target.invoice_number });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
