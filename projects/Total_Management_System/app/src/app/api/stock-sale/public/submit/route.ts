import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { sendAdminEmail } from '@/lib/notification/email';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';

/**
 * 재고판매(LS) 접수 API — 비인증 + CORS (2026-07-21, 단계2)
 *
 * 고객 카탈로그 폼(page.mamoru.kr/projects/stock_sale) 이 POST.
 * EVENT 와 같은 event_submissions 를 쓰되 kind='stock_sale' 로 구분(마이그 117).
 * 입금확인 시 판매전환은 EVENT 와 동일 파이프라인(convertEventToSale) 재사용.
 *
 * ⚠️ 가격은 클라이언트 값을 믿지 않고 서버에서 products 로 재조회한다(위변조 방지).
 * ⚠️ 재고 수량은 여기서 차감하지 않는다 — 입금확인 → 판매전환 시점에 차감(EVENT 와 동일).
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

interface InItem { product_id?: string; qty?: number }

/** 배송비 정책 — 상품금액 5만원 미만이면 3,000원, 이상이면 무료 */
const SHIP_FEE = 3000;
const FREE_SHIP_THRESHOLD = 50000;

/** SS-YYYYMMDD-NNN 자동 채번 (EVENT 의 EV- 와 시퀀스 분리) */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function generateStockNumber(db: any): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `SS-${today}-`;
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

export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { name, phone, postcode, address1, address2, memo, items } = body as {
      name?: string; phone?: string; postcode?: string;
      address1?: string; address2?: string; memo?: string; items?: InItem[];
    };

    if (!name?.trim() || !phone?.trim()) {
      return NextResponse.json({ ok: false, error: '이름과 연락처는 필수입니다' }, { status: 400, headers: CORS_HEADERS });
    }
    if (!Array.isArray(items) || items.length === 0) {
      return NextResponse.json({ ok: false, error: '품목을 1개 이상 선택해주세요' }, { status: 400, headers: CORS_HEADERS });
    }
    if (!address1?.trim()) {
      return NextResponse.json({ ok: false, error: '배송지가 필요합니다' }, { status: 400, headers: CORS_HEADERS });
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // ── 서버에서 가격/재고 재조회 (클라이언트 값 불신) ──
    const ids = [...new Set(items.map((it) => it.product_id).filter(Boolean))] as string[];
    if (ids.length === 0) {
      return NextResponse.json({ ok: false, error: '유효한 품목이 없습니다' }, { status: 400, headers: CORS_HEADERS });
    }
    const { data: prods, error: prodErr } = await dbAny
      .from('products')
      .select('id, name, sku, price, stock_quantity, is_active, category')
      .in('id', ids)
      .eq('category', 'LS')
      .eq('is_active', true);
    if (prodErr) throw prodErr;
    const prodMap = new Map<string, { id: string; name: string; sku: string; price: number; stock_quantity: number }>(
      (prods || []).map((p: { id: string }) => [p.id, p as never]),
    );

    // 서버 권위 품목 라인 (재고 상한 클램프, 품절/미존재 제외)
    const built: { product_id: string; product_name: string; sku: string | null; qty: number; unit_price: number; category: string }[] = [];
    for (const it of items) {
      const p = it.product_id ? prodMap.get(it.product_id) : undefined;
      if (!p) continue;                                   // LS 아님/비활성/삭제 → 제외
      const stock = p.stock_quantity ?? 0;
      const qty = Math.max(1, Math.min(Number(it.qty) || 1, stock > 0 ? stock : 0));
      if (qty <= 0) continue;                             // 품절 → 제외
      built.push({ product_id: p.id, product_name: p.name, sku: p.sku || null, qty, unit_price: p.price ?? 0, category: 'LS' });
    }
    if (built.length === 0) {
      return NextResponse.json({ ok: false, error: '선택하신 품목이 품절되었거나 판매 종료되었습니다' }, { status: 409, headers: CORS_HEADERS });
    }
    const productSum = built.reduce((s, b) => s + b.unit_price * b.qty, 0);
    // 배송비: 상품금액 5만원 미만 3,000원, 이상 무료 (고객 폼과 동일 규칙)
    const shipping = productSum >= FREE_SHIP_THRESHOLD ? 0 : SHIP_FEE;
    const totalAmount = productSum + shipping;

    const phoneNorm = phone.replace(/\D/g, '');
    const orderNumber = await generateStockNumber(dbAny);

    // 고객 자동 생성/병합 (전화번호 기준) + 배송지 최신화
    const { customerId } = await matchOrCreateCustomer(dbAny, {
      phone: phone.trim(),
      name: name.trim(),
      source: 'stock_sale',
      extra: {
        addressRoad: address1 || null,
        addressDetail: address2 || null,
        postcode: postcode || null,
      },
    });

    const insertData = {
      event_number: orderNumber,
      kind: 'stock_sale',
      campaign_id: null,
      customer_id: customerId,
      customer_name: name.trim(),
      customer_phone: phone.trim(),
      receive_method: 'delivery',
      postcode: postcode || null,
      address1: address1 || null,
      address2: address2 || null,
      items: built,
      slicing_addon: 0,
      total_amount: totalAmount,
      status: 'received',
      memo: memo?.trim() || null,
    };

    const { data: order, error: insErr } = await dbAny
      .from('event_submissions').insert(insertData).select().single();
    if (insErr) throw insErr;

    await dbAny.from('event_history').insert({ event_id: order.id, to_status: 'received', note: '재고판매 접수' });

    // 접수확인 + 입금안내 알림톡 (계좌·금액은 솔라피 템플릿에서 렌더)
    try {
      const itemSummary = built.map((b) => `${b.product_name} ${b.qty}개`).join('\n');
      await sendNotification({
        template: 'stock_received',
        phone: phoneNorm,
        name: name.trim(),
        data: {
          id: orderNumber,
          order_number: orderNumber,
          items: itemSummary,
          product_amount: String(productSum),
          shipping_fee: shipping === 0 ? '무료' : String(shipping),
          total_amount: String(totalAmount),
          address: [address1, address2].filter(Boolean).join(' '),
        },
      });
    } catch (notifyErr) {
      console.error('[stock-sale/submit] 알림톡 실패 (접수는 완료):', notifyErr);
    }

    // 관리자 메일
    try {
      const lines = [
        `■ 재고판매 접수`, ``,
        `접수번호: ${orderNumber}`,
        `고객명: ${name.trim()}`,
        `연락처: ${phone}`,
        `품목: ${built.map((b) => `${b.product_name} ${b.qty}개`).join(', ')}`,
        `상품금액: ${productSum.toLocaleString()}원`,
        `배송비: ${shipping === 0 ? '무료' : shipping.toLocaleString() + '원'}`,
        `합계: ${totalAmount.toLocaleString()}원`,
        `주소: ${[address1, address2].filter(Boolean).join(' ')}`,
      ];
      if (memo) lines.push(`메모: ${memo}`);
      await sendAdminEmail(`[MAMORU 재고판매] 새 접수 — ${orderNumber}`, lines.join('\n'));
    } catch (emailErr) {
      console.error('[stock-sale/submit] 이메일 실패:', emailErr);
    }

    return NextResponse.json(
      { ok: true, data: { order_number: orderNumber, total_amount: totalAmount } },
      { headers: CORS_HEADERS },
    );
  } catch (err) {
    console.error('[stock-sale/public/submit] 접수 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
