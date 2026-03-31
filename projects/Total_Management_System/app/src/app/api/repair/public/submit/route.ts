import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** AS-YYYYMMDD-NNN 자동 채번 */
async function generateAsId(db: ReturnType<typeof createServiceClient>): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `AS-${today}-`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data } = await (db as any)
    .from('repairs')
    .select('as_id')
    .like('as_id', `${prefix}%`)
    .order('as_id', { ascending: false })
    .limit(1);

  let seq = 1;
  if (data && data.length > 0) {
    const last = data[0].as_id as string;
    seq = parseInt(last.split('-').pop() || '0', 10) + 1;
  }

  return `${prefix}${String(seq).padStart(3, '0')}`;
}

/** 비용 자동 계산 (GAS Code.js 로직 이전) */
function calculateCosts(qtyMamoru: number, qtyOther: number, proceedType: string) {
  // 수리 비용: 마모루 1만원, 타사 2만원
  const serviceCost = (qtyMamoru * 10000) + (qtyOther * 20000);
  const totalQty = qtyMamoru + qtyOther;

  // 수거비 계산
  let shippingFee = 0;
  if (proceedType === '방문수거') {
    if (totalQty === 1) shippingFee = 6000;
    else if (totalQty === 2) shippingFee = 3000;
    // 3+ : 무료
  } else { // 직접발송
    if (totalQty === 1) shippingFee = 3000;
    // 2+ : 무료
  }

  return { serviceCost, shippingFee, totalAmount: serviceCost + shippingFee };
}

/** POST /api/repair/public/submit — 복원수리 접수 (비인증, CORS) */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      name, phone, postcode, address1, address2,
      proceed_type, pickup_date, delivery_method,
      qty_mamoru, qty_other, memo,
    } = body;

    // 필수값 검증
    if (!name?.trim() || !phone?.trim()) {
      return NextResponse.json(
        { ok: false, error: '이름과 연락처는 필수입니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const qtyM = parseInt(qty_mamoru) || 0;
    const qtyO = parseInt(qty_other) || 0;
    if (qtyM + qtyO < 1) {
      return NextResponse.json(
        { ok: false, error: '가위 수량을 입력해주세요' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 중복 접수 체크 (같은 전화번호 + 주소, 24시간 이내)
    const phoneNorm = phone.replace(/\D/g, '');
    const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
    const addrNorm = (address1 || '').replace(/[\s\-\.]/g, '').toLowerCase();

    if (addrNorm) {
      const { data: recent } = await dbAny
        .from('repairs')
        .select('as_id')
        .eq('phone_normalized', phoneNorm)
        .gte('created_at', oneDayAgo)
        .neq('status', 'cancelled')
        .limit(5);

      if (recent && recent.length > 0) {
        // 주소 유사도 체크 (정규화 후 비교)
        const { data: recentFull } = await dbAny
          .from('repairs')
          .select('as_id, address1')
          .in('as_id', recent.map((r: { as_id: string }) => r.as_id));

        const hasDup = (recentFull || []).some((r: { address1: string }) => {
          const rAddr = (r.address1 || '').replace(/[\s\-\.]/g, '').toLowerCase();
          return rAddr === addrNorm;
        });

        if (hasDup) {
          return NextResponse.json(
            { ok: false, error: '24시간 이내에 동일한 접수가 있습니다. 중복 접수인지 확인해주세요.' },
            { status: 409, headers: CORS_HEADERS }
          );
        }
      }
    }

    // 비용 계산
    const { serviceCost, shippingFee, totalAmount } = calculateCosts(qtyM, qtyO, proceed_type || '직접발송');

    // AS ID 채번
    const asId = await generateAsId(db);

    // INSERT
    const insertData = {
      as_id: asId,
      name: name.trim(),
      phone: phone.trim(),
      // phone_normalized는 DB generated column — INSERT 제외
      proceed_type: proceed_type || '직접발송',
      postcode: postcode || null,
      address1: address1 || null,
      address2: address2 || null,
      pickup_date: pickup_date || null,
      delivery_method: delivery_method || null,
      qty_mamoru: qtyM,
      qty_other: qtyO,
      memo: memo?.trim() || null,
      service_cost: serviceCost,
      shipping_fee: shippingFee,
      total_amount: totalAmount,
      status: 'intake',
      received_at: new Date().toISOString(),
    };

    const { data: repair, error: insertErr } = await dbAny
      .from('repairs')
      .insert(insertData)
      .select()
      .single();

    if (insertErr) throw insertErr;

    // 상태 이력 기록
    await dbAny.from('repair_history').insert({
      repair_id: repair.id,
      to_status: 'intake',
      note: '고객 접수',
    });

    // 알림톡 발송 (접수 안내)
    try {
      await sendNotification({
        template: 'as_received',
        phone: phoneNorm,
        name: name.trim(),
        data: {
          id: asId,
          as_id: asId,
          qty: String(qtyM + qtyO),
          service_cost: String(serviceCost),
          shipping_fee: String(shippingFee),
          total_amount: String(totalAmount),
          proceed_type: proceed_type || '직접발송',
        },
      });
    } catch (notifyErr) {
      console.error('[repair/submit] 알림톡 발송 실패 (접수는 완료):', notifyErr);
    }

    return NextResponse.json(
      { ok: true, data: { as_id: asId, service_cost: serviceCost, shipping_fee: shippingFee, total_amount: totalAmount } },
      { headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[repair/public/submit] 접수 실패:', err);
    return NextResponse.json(
      { ok: false, error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
