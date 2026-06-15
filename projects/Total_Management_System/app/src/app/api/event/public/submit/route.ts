import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { sendAdminEmail } from '@/lib/notification/email';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';
import { SLICING_ADDON } from '@/lib/event/options';
import type { EventItem } from '@/lib/event/types';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** EV-YYYYMMDD-NNN 자동 채번 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function generateEventNumber(db: any): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `EV-${today}-`;
  const { data } = await db
    .from('event_submissions')
    .select('event_number')
    .like('event_number', `${prefix}%`)
    .order('event_number', { ascending: false })
    .limit(1);
  let seq = 1;
  if (data && data.length > 0) {
    seq = parseInt((data[0].event_number as string).split('-').pop() || '0', 10) + 1;
  }
  return `${prefix}${String(seq).padStart(3, '0')}`;
}

/** POST /api/event/public/submit — EVENT 접수 (비인증, CORS) */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      name, phone, receive_method,
      postcode, address1, address2,
      items, memo, campaign_id,
    } = body as {
      name?: string; phone?: string; receive_method?: string;
      postcode?: string; address1?: string; address2?: string;
      items?: EventItem[]; memo?: string; campaign_id?: string;
    };

    if (!name?.trim() || !phone?.trim()) {
      return NextResponse.json({ ok: false, error: '이름과 연락처는 필수입니다' }, { status: 400, headers: CORS_HEADERS });
    }
    if (!Array.isArray(items) || items.length === 0) {
      return NextResponse.json({ ok: false, error: '품목을 1개 이상 선택해주세요' }, { status: 400, headers: CORS_HEADERS });
    }
    const isVisit = receive_method === 'visit';
    if (!isVisit && !address1?.trim()) {
      return NextResponse.json({ ok: false, error: '택배 발송은 배송지가 필요합니다' }, { status: 400, headers: CORS_HEADERS });
    }

    // 금액 계산 (서버 신뢰 — 단가는 클라이언트 값, 슬라이싱은 서버 가산)
    let baseTotal = 0;
    let slicingAddon = 0;
    for (const it of items) {
      const qty = Math.max(1, parseInt(String(it.qty)) || 1);
      const unit = Math.max(0, parseInt(String(it.unit_price)) || 0);
      baseTotal += unit * qty;
      if (it.slicing) slicingAddon += SLICING_ADDON * qty;
    }
    const totalAmount = baseTotal + slicingAddon;

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    const phoneNorm = phone.replace(/\D/g, '');
    const eventNumber = await generateEventNumber(dbAny);

    // 캠페인 연결 — 지정값 없으면 활성 기본 캠페인
    let campaignId = campaign_id || null;
    if (!campaignId) {
      const { data: def } = await dbAny.from('event_campaigns')
        .select('id').eq('is_default', true).eq('status', 'active').order('created_at').limit(1);
      campaignId = def && def.length > 0 ? def[0].id : null;
    }

    const { customerId } = await matchOrCreateCustomer(dbAny, {
      phone: phone.trim(),
      name: name.trim(),
      source: 'event',
      extra: {
        addressRoad: isVisit ? null : (address1 || null),
        addressDetail: isVisit ? null : (address2 || null),
        postcode: isVisit ? null : (postcode || null),
      },
    });

    const insertData = {
      event_number: eventNumber,
      campaign_id: campaignId,
      customer_id: customerId,
      customer_name: name.trim(),
      customer_phone: phone.trim(),
      receive_method: isVisit ? 'visit' : 'delivery',
      postcode: isVisit ? null : (postcode || null),
      address1: isVisit ? null : (address1 || null),
      address2: isVisit ? null : (address2 || null),
      items,
      slicing_addon: slicingAddon,
      total_amount: totalAmount,
      status: 'received',
      memo: memo?.trim() || null,
    };

    const { data: ev, error: insertErr } = await dbAny
      .from('event_submissions')
      .insert(insertData)
      .select()
      .single();
    if (insertErr) throw insertErr;

    await dbAny.from('event_history').insert({ event_id: ev.id, to_status: 'received', note: '고객 접수' });

    // 접수확인 알림톡 (자동)
    try {
      const itemSummary = items.map((it) => `${it.product_name}${it.slicing ? '(슬라이싱)' : ''} ${it.qty}개`).join(', ');
      await sendNotification({
        template: 'event_received',
        phone: phoneNorm,
        name: name.trim(),
        data: {
          id: eventNumber,
          event_number: eventNumber,
          items: itemSummary,
          total_amount: String(totalAmount),
          receive_method: isVisit ? '매장방문' : '택배발송',
        },
      });
    } catch (notifyErr) {
      console.error('[event/submit] 알림톡 실패 (접수는 완료):', notifyErr);
    }

    // 관리자 메일
    try {
      const lines = [
        `■ EVENT 접수 알림`, ``,
        `접수번호: ${eventNumber}`,
        `고객명: ${name.trim()}`,
        `연락처: ${phone}`,
        `수령방법: ${isVisit ? '매장방문' : '택배발송'}`,
        `품목: ${items.map((it) => `${it.product_name}${it.slicing ? '(슬라이싱)' : ''} ${it.qty}개`).join(', ')}`,
        `합계: ${totalAmount.toLocaleString()}원`,
      ];
      if (!isVisit) lines.push(`주소: ${[address1, address2].filter(Boolean).join(' ')}`);
      if (memo) lines.push(`메모: ${memo}`);
      await sendAdminEmail(`[MAMORU EVENT] 새 접수 — ${eventNumber}`, lines.join('\n'));
    } catch (emailErr) {
      console.error('[event/submit] 이메일 실패:', emailErr);
    }

    return NextResponse.json(
      { ok: true, data: { event_number: eventNumber, total_amount: totalAmount } },
      { headers: CORS_HEADERS },
    );
  } catch (err) {
    console.error('[event/public/submit] 접수 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
