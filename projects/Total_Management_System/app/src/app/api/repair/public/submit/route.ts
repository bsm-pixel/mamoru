import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { sendAdminEmail } from '@/lib/notification/email';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';
import { fireAndForgetRepairSync } from '@/lib/google/repair-calendar-sync';

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
  if (proceedType === '직접방문') {
    // 2026-05-25: 매장 직접방문(당일수리) — 배송 없음
    shippingFee = 0;
  } else if (proceedType === '방문수거') {
    if (totalQty === 1) shippingFee = 6000;
    else if (totalQty === 2) shippingFee = 3000;
    // 3+ : 무료
  } else { // 직접발송
    if (totalQty === 1) shippingFee = 3000;
    // 2+ : 무료
  }

  return { serviceCost, shippingFee, totalAmount: serviceCost + shippingFee };
}

/** 'YYYY-MM-DD' → 'YYYY년 MM월 DD일 (요일)'. 서버(UTC) 타임존 무관 — Date.UTC+getUTCDay로 KST 날짜 그대로 표기
 *  (이전: `new Date(date+'T00:00:00+09:00').getDay()` → UTC 서버에서 하루 전 요일/날짜로 밀리는 버그) */
function formatKoreanDate(dateStr: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return '';
  const [y, m, d] = dateStr.split('-').map(Number);
  const dow = ['일', '월', '화', '수', '목', '금', '토'][new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
  return `${y}년 ${String(m).padStart(2, '0')}월 ${String(d).padStart(2, '0')}일 (${dow}요일)`;
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

    // 비용 계산
    const { serviceCost, shippingFee, totalAmount } = calculateCosts(qtyM, qtyO, proceed_type || '직접발송');

    // AS ID 채번
    const asId = await generateAsId(db);

    // 고객 자동 매칭/생성 — phone 기준 SSOT
    const { customerId } = await matchOrCreateCustomer(dbAny, {
      phone: phone.trim(),
      name: name.trim(),
      source: 'as',
      extra: {
        addressRoad: address || null,
        addressDetail: address_detail || null,
        postcode: postcode || null,
      },
    });

    // 직접방문 차단 시간 서버 계산 (클라이언트 값 신뢰 X — 충돌 검사 정합성)
    // 소요시간 = 10분 + 자루당 5분 (slots API 와 동일 공식, 2026-05-27 사장님 공식)
    let visitDuration: number | null = null;
    if (proceed_type === '직접방문') {
      visitDuration = 10 + (Math.max(qtyM + qtyO, 1) - 1) * 5;
    }

    const isVisit = proceed_type === '직접방문';

    // INSERT
    const insertData = {
      customer_id: customerId,
      as_id: asId,
      name: name.trim(),
      phone: phone.trim(),
      // phone_normalized는 DB generated column — INSERT 제외
      proceed_type: proceed_type || '직접발송',
      postcode: isVisit ? null : (postcode || null),
      address: isVisit ? null : (address || null),
      address_detail: isVisit ? null : (address_detail || null),
      pickup_date: isVisit ? null : (pickup_date || null),
      delivery_method: isVisit ? null : (delivery_method || null),
      // 2026-05-25: 직접방문(당일수리) 신규 컬럼
      visit_date: isVisit ? visit_date : null,
      visit_time: isVisit ? visit_time : null,
      visit_duration_min: visitDuration,
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

    // 2026-05-25 Phase 3-B: 직접방문 → Google Calendar 자동 동기화 (fire-and-forget)
    if (isVisit) {
      after(() => fireAndForgetRepairSync(repair.id));
    }

    // pickup_date 표시 포맷 (방문수거 알림톡) — 타임존 무관
    const pickupDateDisplay = pickup_date ? formatKoreanDate(pickup_date) : '';
    // 방문일 표시 포맷 (직접방문 알림톡)
    const visitDateDisplay = (isVisit && visit_date) ? formatKoreanDate(visit_date) : '';

    // 알림톡 발송 (접수 안내)
    // 직접방문은 별도 템플릿 'as_visit_booked' Phase 4 검수 후 활성화 — 현재는 as_received 임시 사용
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
          delivery_method: delivery_method || '',
          pickup_date: pickupDateDisplay,
          postcode: postcode || '',
          address: address || '',
          address_detail: address_detail || '',
          pickup_address_text: [address, address_detail].filter(Boolean).join(' '),
          // 직접방문 신규 변수 (Phase 4 템플릿 신청 시 활용)
          visit_date: visitDateDisplay,
          visit_time: isVisit ? (visit_time || '') : '',
          visit_duration_min: visitDuration ? String(visitDuration) : '',
        },
      });
    } catch (notifyErr) {
      console.error('[repair/submit] 알림톡 발송 실패 (접수는 완료):', notifyErr);
    }

    // Gmail 관리자 알림 (상담 submit과 동일 패턴)
    try {
      const emailLines = [
        `■ 복원수리 접수 알림`,
        ``,
        `접수번호: ${asId}`,
        `고객명: ${name.trim()}`,
        `연락처: ${phone}`,
        `진행방식: ${proceed_type || '직접발송'}`,
        `마모루: ${qty_mamoru || 0}정, 타사: ${qty_other || 0}정`,
        `주소: ${[address, address_detail].filter(Boolean).join(' ')}`,
      ];
      if (memo) emailLines.push(`메모: ${memo}`);
      await sendAdminEmail(`[MAMORU 복원수리] 새 접수 — ${asId}`, emailLines.join('\n'));
    } catch (emailErr) {
      console.error('[repair/submit] 이메일 발송 실패:', emailErr);
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
