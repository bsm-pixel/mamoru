import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { findPrefillAddress, upsertCustomerAddress, digits, type ShipmentRow } from '@/lib/reviews/event-shipments';

/**
 * 리뷰 이벤트 당첨자 배송지 입력 — 고객 공개 API (2026-10-02, 마이그 156)
 *   GET  ?t=<token>  → 이름·연락처(고정) + 저장된 주소 또는 예전에 남긴 주소(미리 채움)
 *   POST { t, postcode, address_road, address_detail, delivery_message } → 저장 + 고객정보 주소 갱신
 *
 * 🔓 로그인 없음 — 알림톡 링크의 토큰(추측 불가 32자)이 곧 권한. 토큰 없는 요청은 아무것도 돌려주지 않는다.
 * 페이지: projects/reviews/page_event_address.html
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const TOKEN_RE = /^[0-9a-f]{32}$/;

export async function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

const SOURCE_LABEL: Record<string, string> = {
  customer: '이전에 남겨주신 주소',
  repair: '복원수리 접수 당시 기재하신 주소',
  order: '제품 주문 당시 기재하신 주소',
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function loadRow(db: any, token: string): Promise<ShipmentRow | null> {
  if (!TOKEN_RE.test(token)) return null;
  const { data } = await db.from('review_event_shipments').select('*').eq('token', token).is('cancelled_at', null).maybeSingle();
  return data || null;
}

function phoneFmt(p: string): string {
  const d = digits(p);
  return d.length === 11 ? `${d.slice(0, 3)}-${d.slice(3, 7)}-${d.slice(7)}` : d;
}

export async function GET(req: NextRequest) {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const row = await loadRow(db, req.nextUrl.searchParams.get('t') || '');
    if (!row) return NextResponse.json({ error: 'not_found' }, { status: 404, headers: CORS_HEADERS });

    let address = row.address_road
      ? { postcode: row.postcode || '', address_road: row.address_road, address_detail: row.address_detail || '' }
      : null;
    let source: string | null = row.address_submitted_at ? 'saved' : null;

    if (!address) {
      const pre = await findPrefillAddress(db, row.phone);
      if (pre) {
        address = { postcode: pre.postcode, address_road: pre.address_road, address_detail: pre.address_detail };
        source = pre.source;
        if (!row.address_prefill_source) {
          await db.from('review_event_shipments').update({ address_prefill_source: pre.source }).eq('id', row.id);
        }
      }
    }

    return NextResponse.json({
      name: row.name,
      phone: phoneFmt(row.phone),
      rank_label: row.rank_label,
      prize: row.prize,
      address,
      delivery_message: row.delivery_message || '',
      source,                                         // saved | customer | repair | order | null
      source_label: source && source !== 'saved' ? SOURCE_LABEL[source] : null,
      submitted_at: row.address_submitted_at,
      locked: !!row.invoice_number,                   // 송장 생성 후엔 수정 불가 (발송 준비 중)
      shipped: !!row.shipped_at,
    }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json().catch(() => ({}));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const row = await loadRow(db, String(body.t || ''));
    if (!row) return NextResponse.json({ error: 'not_found' }, { status: 404, headers: CORS_HEADERS });
    if (row.invoice_number) {
      return NextResponse.json({ error: '이미 발송 준비 중이라 주소를 바꿀 수 없습니다. 고객센터로 연락 주세요' }, { status: 409, headers: CORS_HEADERS });
    }

    const postcode = String(body.postcode || '').trim().slice(0, 10);
    const address_road = String(body.address_road || '').trim().slice(0, 200);
    const address_detail = String(body.address_detail || '').trim().slice(0, 100);
    const delivery_message = String(body.delivery_message || '').trim().slice(0, 60);
    if (!postcode || !address_road) {
      return NextResponse.json({ error: '주소 검색으로 주소를 선택해 주세요' }, { status: 400, headers: CORS_HEADERS });
    }
    if (!address_detail) {
      return NextResponse.json({ error: '상세 주소(동·호수 등)를 입력해 주세요' }, { status: 400, headers: CORS_HEADERS });
    }

    const firstTime = !row.address_submitted_at;
    // 고객정보 주소도 새 주소로 (사장님 결정 2026-10-02)
    const customerId = await upsertCustomerAddress(db, { name: row.name, phone: row.phone, postcode, address_road, address_detail });
    const now = new Date().toISOString();
    await db.from('review_event_shipments').update({
      postcode, address_road, address_detail, delivery_message: delivery_message || null,
      address_submitted_at: now, customer_id: customerId, updated_at: now,
    }).eq('id', row.id);

    // 고객 행동 → 사장님 푸시 (주소가 들어와야 송장을 만들 수 있으므로)
    const pushTitle = firstTime ? '당첨자 배송지 입력 📦' : '당첨자 배송지 수정';
    const deliverPush = async () => {
      const { sendPushToAll } = await import('@/lib/firebase/send-push');
      await sendPushToAll({
        title: pushTitle,
        body: `${row.name}님 ${row.rank_label || ''} — 송장 생성 가능`,
        url: '/reviews/event',
        tag: `mamoru-event-address-${row.id}-${Date.now()}`,
      });
    };
    try { after(deliverPush); } catch { await deliverPush().catch(() => {}); }

    return NextResponse.json({ ok: true }, { headers: CORS_HEADERS });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
