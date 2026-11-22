import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { errMsg } from '@/lib/utils/err';
import { CAUTION_NOTICE, LEGAL_NOTICE, ENGRAVING_NOTICE } from '@/lib/contracts/notice';
import { cleanEngraving, isToken, isValidSignature, maskName } from '@/lib/contracts/sign-link';

/**
 * 고객 서명 링크 — 로그인 없이 열쇠(?t=sign_token)로만 (2026-11-22, 마이그 172)
 *  화면: projects/contract/page_sign.html (page.mamoru.kr)
 *
 *  GET  ?t=  → 계약 내용(제품·금액·결제·유의사항). 이름은 성만 가린 꼴, 전화·주소는 보내지 않는다.
 *  POST { t, action:'sign', signature, agreed, engraving_agreed?, engraving_text? } → 서명 (한 번만)
 *  POST { t, action:'engraving', text } → 「나중에 받기」였던 각인 문구 보내기 (받기 전까지만)
 *  서명·문구가 들어오면 사장님 휴대폰 푸시. 고객 알림톡은 보내지 않는다.
 */

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};
const json = (body: unknown, status = 200) => NextResponse.json(body, { status, headers: CORS });

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS });
}

const PAYMENT_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '현금 · 계좌이체', mixed: '복합', cms: 'CMS 자동이체',
};
const ACCOUNT = '우리은행 1002-439-462514 (백성민)';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function findByToken(db: any, t: string) {
  const { data } = await db.from('contracts')
    .select('id, contract_number, customer_name, status, total_amount, discount_amount, final_amount, payment_method, installment_months, deposit_amount, balance_amount, signature_data, signed_at, created_at, engraving, engraving_text, engraving_status')
    .eq('sign_token', t).maybeSingle();
  return data;
}

export async function GET(req: NextRequest) {
  try {
    const t = req.nextUrl.searchParams.get('t');
    if (!isToken(t)) return json({ ok: false, error: '잘못된 주소입니다. 받으신 링크로 다시 들어와 주세요.' }, 400);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const c = await findByToken(db, t);
    if (!c) return json({ ok: false, error: '계약서를 찾을 수 없어요. 매장으로 문의해 주세요.' }, 404);
    const { data: items } = await db.from('contract_items').select('product_name, quantity, unit_price, total_price').eq('contract_id', c.id);

    return json({
      ok: true,
      data: {
        number: c.contract_number,
        name: maskName(c.customer_name),
        status: c.status,
        createdAt: c.created_at,
        signed: !!c.signature_data,
        signedAt: c.signed_at,
        items: (items || []).map((i: { product_name: string; quantity: number; unit_price: number; total_price: number }) => ({ name: i.product_name, qty: i.quantity, price: i.unit_price, total: i.total_price })),
        total: c.total_amount, discount: c.discount_amount || 0, final: c.final_amount,
        deposit: c.deposit_amount || 0, balance: c.balance_amount || 0,
        payment: PAYMENT_LABEL[c.payment_method] || c.payment_method,
        installment: c.installment_months || 0,
        account: ACCOUNT,
        engraving: !!c.engraving,
        engravingText: c.engraving_text || '',
        engravingStatus: c.engraving_status || null,
        notices: { legal: LEGAL_NOTICE, caution: CAUTION_NOTICE, engraving: ENGRAVING_NOTICE },
      },
    });
  } catch (err) {
    return json({ ok: false, error: errMsg(err) }, 500);
  }
}

export async function POST(req: NextRequest) {
  try {
    const body = await req.json().catch(() => ({}));
    const t = body?.t;
    if (!isToken(t)) return json({ ok: false, error: '잘못된 주소입니다.' }, 400);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const c = await findByToken(db, t);
    if (!c) return json({ ok: false, error: '계약서를 찾을 수 없어요.' }, 404);
    if (c.status === 'cancelled') return json({ ok: false, error: '취소된 계약서예요. 매장으로 문의해 주세요.' }, 409);
    const now = new Date().toISOString();

    if (body.action === 'sign') {
      if (c.signature_data) return json({ ok: false, error: '이미 서명하신 계약서예요.' }, 409);
      if (body.agreed !== true) return json({ ok: false, error: '유의사항에 동의해 주세요.' }, 400);
      if (c.engraving && body.engraving_agreed !== true) return json({ ok: false, error: '이니셜 각인 안내에 동의해 주세요.' }, 400);
      if (!isValidSignature(body.signature)) return json({ ok: false, error: '서명을 다시 해 주세요.' }, 400);

      const patch: Record<string, unknown> = {
        signature_data: body.signature, signed_at: now, signed_via: 'link', terms_agreed_at: now, status: 'signed',
      };
      const eng = cleanEngraving(body.engraving_text);
      if (c.engraving && c.engraving_status === 'pending' && eng) Object.assign(patch, { engraving_text: eng, engraving_status: 'received', engraving_at: now });

      // 한 번만 — 그 사이 다른 곳에서 서명됐으면 덮어쓰지 않는다
      const { data: done, error } = await db.from('contracts').update(patch)
        .eq('id', c.id).is('signature_data', null).select('id').maybeSingle();
      if (error) throw error;
      if (!done) return json({ ok: false, error: '이미 서명하신 계약서예요.' }, 409);

      const engLine = c.engraving ? (patch.engraving_text ? ` · 각인 「${patch.engraving_text}」` : ' · 각인 문구 대기') : '';
      after(async () => {
        try {
          const { sendPushToAll } = await import('@/lib/firebase/send-push');
          await sendPushToAll({ title: '✍️ 계약서 서명 완료', body: `${c.contract_number} ${c.customer_name}님이 링크로 서명했어요${engLine}`, url: `/contracts/${c.id}`, tag: `contract-sign-${c.id}` });
        } catch (e) { console.error('[contract sign] 푸시 실패', e); }
      });
      return json({ ok: true, engravingPending: !!c.engraving && !patch.engraving_text });
    }

    if (body.action === 'engraving') {
      if (!c.engraving) return json({ ok: false, error: '이 계약서에는 각인이 없어요.' }, 400);
      if (c.engraving_status !== 'pending') return json({ ok: false, error: '각인 문구를 이미 받았어요. 바꾸시려면 매장으로 말씀해 주세요.' }, 409);
      const text = cleanEngraving(body.text);
      if (!text) return json({ ok: false, error: '각인할 문구를 적어 주세요.' }, 400);
      const { data: done, error } = await db.from('contracts')
        .update({ engraving_text: text, engraving_status: 'received', engraving_at: now })
        .eq('id', c.id).eq('engraving_status', 'pending').select('id').maybeSingle();
      if (error) throw error;
      if (!done) return json({ ok: false, error: '각인 문구를 이미 받았어요.' }, 409);
      after(async () => {
        try {
          const { sendPushToAll } = await import('@/lib/firebase/send-push');
          await sendPushToAll({ title: '🔤 각인 문구 도착', body: `${c.contract_number} ${c.customer_name}님 — 「${text}」`, url: `/contracts/${c.id}`, tag: `contract-eng-${c.id}` });
        } catch (e) { console.error('[contract engraving] 푸시 실패', e); }
      });
      return json({ ok: true, text });
    }

    return json({ ok: false, error: '알 수 없는 요청' }, 400);
  } catch (err) {
    return json({ ok: false, error: errMsg(err) }, 500);
  }
}
