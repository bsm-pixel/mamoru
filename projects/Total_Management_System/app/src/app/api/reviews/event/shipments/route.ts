import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { getNextInvoice, bookShipment, cancelShipment } from '@/lib/lotte/alps-client';
import {
  syncShipments, findPrefillAddress, upsertCustomerAddress, sendWonNotice, addressLink, digits,
  type ShipmentRow,
} from '@/lib/reviews/event-shipments';

/**
 * 리뷰 이벤트 당첨자 배송 — 관리 API (2026-10-02, 마이그 156)
 *   GET  ?month=YYMM                → 당첨자 기준 동기화 후 배송 목록 (+예전 주소 미리보기)
 *   POST { action: 'notify', month, ids? }           → 당첨 안내 알림톡 (ids 없으면 아직 안 보낸 전원)
 *   POST { action: 'invoice', id }                   → 롯데 ALPS 송장 생성
 *   POST { action: 'cancel_invoice', id }            → 송장 취소 (집하 전만)
 *   POST { action: 'save_address', id, postcode, address_road, address_detail, delivery_message }
 *                                                    → 사장님이 전화로 받은 주소 직접 입력
 */

async function auth() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  return user;
}

export async function GET(req: NextRequest) {
  try {
    if (!(await auth())) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const month = req.nextUrl.searchParams.get('month') || '';
    if (!/^\d{4}$/.test(month)) return NextResponse.json({ error: 'month(YYMM) required' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const rows = await syncShipments(db, month);

    // 주소 미입력자는 "예전에 남긴 주소가 있는지" 미리 보여준다 (사장님 판단용)
    const items = await Promise.all(rows.map(async (r) => {
      const prefill = r.address_road ? null : await findPrefillAddress(db, r.phone);
      return { ...r, address_link: addressLink(r.token), prefill };
    }));
    return NextResponse.json({ items });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    if (!(await auth())) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const body = await req.json().catch(() => ({}));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    // ── 당첨 안내 알림톡 ──
    if (body.action === 'notify') {
      const month = String(body.month || '');
      const rows = await syncShipments(db, month);
      const ids: string[] | null = Array.isArray(body.ids) && body.ids.length ? body.ids : null;
      const targets = rows.filter((r) => (ids ? ids.includes(r.id) : !r.won_notified_at));
      let sent = 0;
      const failed: string[] = [];
      for (const r of targets) {
        if (!digits(r.phone)) { failed.push(`${r.name}(연락처 없음)`); continue; }
        const res = await sendWonNotice(db, r);
        if (res.ok) sent++; else failed.push(`${r.name}(${res.error || '실패'})`);
      }
      return NextResponse.json({ ok: true, sent, failed, total: targets.length });
    }

    const id = String(body.id || '');
    const { data: row } = await db.from('review_event_shipments').select('*').eq('id', id).maybeSingle() as { data: ShipmentRow | null };
    if (!row) return NextResponse.json({ error: '당첨자 배송 정보를 찾을 수 없습니다' }, { status: 404 });

    // ── 주소 직접 입력 (전화로 받은 경우) ──
    if (body.action === 'save_address') {
      const postcode = String(body.postcode || '').trim();
      const address_road = String(body.address_road || '').trim();
      const address_detail = String(body.address_detail || '').trim();
      if (!postcode || !address_road) return NextResponse.json({ error: '우편번호와 주소를 입력해주세요' }, { status: 400 });
      if (row.invoice_number) return NextResponse.json({ error: '이미 송장이 생성되어 주소를 바꿀 수 없습니다. 송장 취소 후 수정하세요' }, { status: 409 });
      const customerId = await upsertCustomerAddress(db, { name: row.name, phone: row.phone, postcode, address_road, address_detail });
      const now = new Date().toISOString();
      await db.from('review_event_shipments').update({
        postcode, address_road, address_detail: address_detail || null,
        delivery_message: String(body.delivery_message || '').trim() || null,
        address_submitted_at: now, customer_id: customerId, updated_at: now,
      }).eq('id', id);
      return NextResponse.json({ ok: true });
    }

    // ── 송장 생성 ──
    if (body.action === 'invoice') {
      if (row.invoice_number) return NextResponse.json({ error: `이미 송장이 있습니다 (${row.invoice_number})` }, { status: 409 });
      if (!row.postcode || !row.address_road) return NextResponse.json({ error: '배송지가 아직 없습니다. 고객 입력을 기다리거나 직접 입력해주세요' }, { status: 400 });
      if (!digits(row.phone)) return NextResponse.json({ error: '연락처가 없습니다' }, { status: 400 });

      const { invoiceNumber } = await getNextInvoice();
      const goodsName = `리뷰이벤트 ${row.prize || row.rank_label || '당첨상품'}`.slice(0, 50);
      const result = await bookShipment({
        invoiceNumber,
        receiverName: row.name,
        receiverTel: digits(row.phone),
        receiverZip: row.postcode,
        receiverAddr: [row.address_road, row.address_detail].filter(Boolean).join(' '),
        goodsName,
        deliveryMessage: row.delivery_message ?? undefined,
      });
      if (!result.success) return NextResponse.json({ error: `ALPS 송장 생성 실패: ${result.error}` }, { status: 502 });

      // 동시 클릭 방지 — 송장이 비어 있을 때만 기록
      const { data: saved } = await db.from('review_event_shipments')
        .update({ invoice_number: invoiceNumber, invoice_created_at: new Date().toISOString(), courier_name: '롯데택배', updated_at: new Date().toISOString() })
        .eq('id', id).is('invoice_number', null).select('id');
      if (!saved?.length) {
        console.error('[event-shipments] 송장 중복 발급 — ALPS 취소 시도:', invoiceNumber);
        await cancelShipment(invoiceNumber).catch(() => {});
        return NextResponse.json({ error: '다른 화면에서 이미 송장이 생성되었습니다. 새로고침 해주세요' }, { status: 409 });
      }
      return NextResponse.json({ ok: true, invoiceNumber });
    }

    // ── 송장 취소 (집하 전만) ──
    if (body.action === 'cancel_invoice') {
      if (!row.invoice_number) return NextResponse.json({ error: '취소할 송장이 없습니다' }, { status: 400 });
      if (row.shipped_at) return NextResponse.json({ error: '이미 집하(출고)된 송장은 취소할 수 없습니다' }, { status: 409 });
      const r = await cancelShipment(row.invoice_number);
      if (!r.success) return NextResponse.json({ error: `ALPS 송장 취소 실패: ${r.error} — ALPS에서 직접 취소 후 다시 시도하세요` }, { status: 502 });
      await db.from('review_event_shipments')
        .update({ invoice_number: null, invoice_created_at: null, updated_at: new Date().toISOString(), memo: `송장 ${row.invoice_number} 취소 (${new Date().toISOString().slice(0, 10)})` })
        .eq('id', id);
      return NextResponse.json({ ok: true });
    }

    return NextResponse.json({ error: 'unknown action' }, { status: 400 });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
