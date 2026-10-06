/**
 * 복원수리 접수 생성 — 공용 (2026-10-06)
 *
 * 고객 접수 페이지(api/repair/public/submit)와 TMS 「방문 예약 등록」(api/repair/visit-booking)이
 * 같은 함수를 쓴다 → 예약번호·비용·알림톡·캘린더가 두 갈래로 갈라지지 않는다.
 * (원래 public/submit 안에 있던 로직을 그대로 옮김 — 동작 변경 없음. 차이는 actor 만)
 *
 *   actor 'customer' : 이력 "고객 접수", 사장님 앱 푸시 O, 관리자 메일 O
 *   actor 'admin'    : 이력 "관리자 전화 예약 등록", 사장님 앱 푸시 X(본인 행동), 관리자 메일 X
 *   고객 알림톡은 둘 다 항상 발송 (알림톡 항상 발송 원칙)
 */

import { after } from 'next/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { sendAdminEmail } from '@/lib/notification/email';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';
import { syncRepairToCalendar } from '@/lib/google/repair-calendar-sync';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

/** AS-YYYYMMDD-NNN 자동 채번 */
export async function generateAsId(db: Db): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `AS-${today}-`;
  const { data } = await db
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
export function calculateCosts(qtyMamoru: number, qtyOther: number, proceedType: string) {
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

/** 'YYYY-MM-DD' → 'YYYY년 MM월 DD일 (요일)'. 서버(UTC) 타임존 무관 — Date.UTC+getUTCDay로 KST 날짜 그대로 표기 */
export function formatKoreanDate(dateStr: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return '';
  const [y, m, d] = dateStr.split('-').map(Number);
  const dow = ['일', '월', '화', '수', '목', '금', '토'][new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
  return `${y}년 ${String(m).padStart(2, '0')}월 ${String(d).padStart(2, '0')}일 (${dow}요일)`;
}

/** 직접방문 소요시간 = 10분 + 자루당 5분 (slots 와 동일 공식, 2026-05-27 사장님 공식) */
export function visitDurationMin(totalQty: number): number {
  return 10 + (Math.max(totalQty, 1) - 1) * 5;
}

export interface RepairIntakeInput {
  name: string;
  phone: string;
  proceed_type?: string;          // 직접방문 | 방문수거 | 직접발송
  postcode?: string | null;
  address?: string | null;
  address_detail?: string | null;
  pickup_date?: string | null;
  delivery_method?: string | null;
  visit_date?: string | null;     // 직접방문
  visit_time?: string | null;     // 직접방문 HH:MM
  qty_mamoru: number;
  qty_other: number;
  memo?: string | null;
}

export async function createRepairIntake(db: Db, input: RepairIntakeInput, actor: 'customer' | 'admin') {
  const name = input.name.trim();
  const phone = input.phone.trim();
  const phoneNorm = phone.replace(/\D/g, '');
  const proceedType = input.proceed_type || '직접발송';
  const isVisit = proceedType === '직접방문';
  const qtyM = input.qty_mamoru;
  const qtyO = input.qty_other;

  const { serviceCost, shippingFee, totalAmount } = calculateCosts(qtyM, qtyO, proceedType);
  const asId = await generateAsId(db);

  // 고객 자동 매칭/생성 — phone 기준 SSOT
  const { customerId } = await matchOrCreateCustomer(db, {
    phone,
    name,
    source: actor === 'admin' ? 'manual' : 'as',
    extra: {
      addressRoad: input.address || null,
      addressDetail: input.address_detail || null,
      postcode: input.postcode || null,
    },
  });

  // 직접방문 차단 시간 서버 계산 (클라이언트 값 신뢰 X — 충돌 검사 정합성)
  const visitDuration = isVisit ? visitDurationMin(qtyM + qtyO) : null;

  const insertData = {
    customer_id: customerId,
    as_id: asId,
    name,
    phone,
    // phone_normalized는 DB generated column — INSERT 제외
    proceed_type: proceedType,
    postcode: isVisit ? null : (input.postcode || null),
    address: isVisit ? null : (input.address || null),
    address_detail: isVisit ? null : (input.address_detail || null),
    pickup_date: isVisit ? null : (input.pickup_date || null),
    delivery_method: isVisit ? null : (input.delivery_method || null),
    visit_date: isVisit ? input.visit_date : null,
    visit_time: isVisit ? input.visit_time : null,
    visit_duration_min: visitDuration,
    qty_mamoru: qtyM,
    qty_other: qtyO,
    memo: input.memo?.trim() || null,
    service_cost: serviceCost,
    shipping_fee: shippingFee,
    total_amount: totalAmount,
    status: 'intake',
    received_at: new Date().toISOString(),
  };

  const { data: repair, error: insertErr } = await db.from('repairs').insert(insertData).select().single();
  if (insertErr) throw insertErr;

  await db.from('repair_history').insert({
    repair_id: repair.id,
    to_status: 'intake',
    note: actor === 'admin' ? '관리자 전화 예약 등록' : '고객 접수',
  });

  // 직접방문 → Google Calendar 자동 동기화. after 가 Promise 를 await 하도록 실제 async 함수를 넘김
  //   (fire-and-forget 로 넘기면 서버리스 함수 종료 시 요청이 잘려 미기록됨 — 2026-08-04 EPIPE 근본수정)
  if (isVisit) {
    try { after(() => syncRepairToCalendar(repair.id)); } catch { await syncRepairToCalendar(repair.id).catch(() => {}); }
  }

  const pickupDateDisplay = input.pickup_date ? formatKoreanDate(input.pickup_date) : '';
  const visitDateDisplay = (isVisit && input.visit_date) ? formatKoreanDate(input.visit_date) : '';

  // 알림톡 (접수 안내) — 직접방문 = as_visit_booked + 일정변경 링크 / 택배·방문수거 = as_received
  //   사장님 본인이 등록한 건(admin)은 사장님 앱 푸시를 보내지 않는다 (본인 행동 = 푸시 제외 규칙)
  let notified = false;
  try {
    const r = isVisit
      ? await sendNotification({
          template: 'as_visit_booked',
          phone: phoneNorm,
          name,
          skipAdminPush: actor === 'admin',
          data: {
            id: asId,
            as_id: asId,
            visit_date: visitDateDisplay,
            visit_time: input.visit_time || '',
            qty: String(qtyM + qtyO),
            visit_duration_min: visitDuration ? String(visitDuration) : '',
            // 일정 확인·변경 버튼 URL (Make 시나리오가 https:// 붙임). manage_token = DB DEFAULT 자동생성
            change_request_link: `page.mamoru.kr/projects/as/page_change_request.html?uid=${repair.manage_token}`,
          },
        })
      : await sendNotification({
          template: 'as_received',
          phone: phoneNorm,
          name,
          skipAdminPush: actor === 'admin',
          data: {
            id: asId,
            as_id: asId,
            qty: String(qtyM + qtyO),
            service_cost: String(serviceCost),
            shipping_fee: String(shippingFee),
            total_amount: String(totalAmount),
            proceed_type: proceedType,
            delivery_method: input.delivery_method || '',
            pickup_date: pickupDateDisplay,
            postcode: input.postcode || '',
            address: input.address || '',
            address_detail: input.address_detail || '',
            pickup_address_text: [input.address, input.address_detail].filter(Boolean).join(' '),
          },
        });
    notified = !!r?.success;
  } catch (notifyErr) {
    console.error('[repair/intake] 알림톡 발송 실패 (접수는 완료):', notifyErr);
  }

  // 관리자 메일 — 고객 접수만 (사장님 본인 등록은 불필요)
  if (actor === 'customer') {
    try {
      const emailLines = [
        `■ 복원수리 접수 알림`,
        ``,
        `접수번호: ${asId}`,
        `고객명: ${name}`,
        `연락처: ${phone}`,
        `진행방식: ${proceedType}`,
        `마모루: ${qtyM}정, 타사: ${qtyO}정`,
        `주소: ${[input.address, input.address_detail].filter(Boolean).join(' ')}`,
      ];
      if (input.memo) emailLines.push(`메모: ${input.memo}`);
      await sendAdminEmail(`[MAMORU 복원수리] 새 접수 — ${asId}`, emailLines.join('\n'));
    } catch (emailErr) {
      console.error('[repair/intake] 이메일 발송 실패:', emailErr);
    }
  }

  return { repair, asId, serviceCost, shippingFee, totalAmount, notified };
}
