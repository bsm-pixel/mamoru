import { NextRequest, NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { createRepairIntake } from '@/lib/repair/intake';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** POST /api/repair/public/submit — 복원수리 접수 (비인증, CORS) */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const {
      name, phone, postcode, address1: address, address2: address_detail,
      proceed_type, pickup_date, delivery_method,
      qty_mamoru, qty_other, memo,
      // 2026-05-25: 직접방문(당일수리) 신규
      visit_date, visit_time,
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

    // 직접방문 입력 검증
    if (proceed_type === '직접방문') {
      if (!visit_date || !/^\d{4}-\d{2}-\d{2}$/.test(visit_date)) {
        return NextResponse.json(
          { ok: false, error: '방문 날짜를 선택해주세요' },
          { status: 400, headers: CORS_HEADERS }
        );
      }
      if (!visit_time || !/^\d{2}:\d{2}$/.test(visit_time)) {
        return NextResponse.json(
          { ok: false, error: '방문 시간을 선택해주세요' },
          { status: 400, headers: CORS_HEADERS }
        );
      }
    }

    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 중복 접수 체크 (같은 전화번호 + 주소, 24시간 이내)
    const phoneNorm = phone.replace(/\D/g, '');
    const oneDayAgo = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
    const addrNorm = (address || '').replace(/[\s\-\.]/g, '').toLowerCase();

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
          .select('as_id, address')
          .in('as_id', recent.map((r: { as_id: string }) => r.as_id));

        const hasDup = (recentFull || []).some((r: { address: string }) => {
          const rAddr = (r.address || '').replace(/[\s\-\.]/g, '').toLowerCase();
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

    // 2026-10-06: 접수 생성(채번·고객연결·INSERT·이력·캘린더·알림톡·관리자메일)은 공용 함수로 — TMS 방문 예약 등록과 같은 경로
    const { asId, serviceCost, shippingFee, totalAmount } = await createRepairIntake(dbAny, {
      name, phone, proceed_type,
      postcode, address, address_detail, pickup_date, delivery_method,
      visit_date, visit_time,
      qty_mamoru: qtyM, qty_other: qtyO, memo,
    }, 'customer');

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
