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

/**
 * 접수번호 AS-YYYYMMDD-NNNN 자동 채번 (2026-10-28 변경)
 *   날짜 = KST (UTC 로 자르면 오전 9시 전 접수가 전날로 찍혔다)
 *   뒤 4자리 = 임의(1000~9999) — 순번(001·002)이면 그날 몇 번째인지 드러나고,
 *   수리내역서 링크(uid=접수번호)에서 끝자리만 바꿔 남의 내역서를 열 수 있었다.
 *   옛 번호(AS-…-001)는 그대로 유효. as_id UNIQUE(repairs_as_id_key)라 겹치면 다시 뽑는다.
 */
export async function generateAsId(db: Db): Promise<string> {
  const today = new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10).replace(/-/g, '');
  for (let i = 0; i < 20; i++) {
    const asId = `AS-${today}-${1000 + Math.floor(Math.random() * 9000)}`;
    const { data } = await db.from('repairs').select('id').eq('as_id', asId).limit(1);
    if (!data || data.length === 0) return asId;
  }
  throw new Error('접수번호 생성 실패 — 다시 시도해 주세요');
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
    // 2026-10-08: 사장님이 「방문 예약 등록」으로 직접 넣은 건은 이미 고객과 확정한 것 → 접수확인까지 된 상태로 저장.
    //   (전엔 confirmed_at 이 비어 '신규접수' 탭에 들어가서, 방문예정 탭에 보이려면 [접수확인]을 한 번 더 눌러야 했다)
    //   고객이 직접 접수한 건은 지금처럼 신규접수 → 사장님 [접수확인] 후 방문예정.
    ...(actor === 'admin' ? { confirmed_at: new Date().toISOString() } : {}),
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

/**
 * 165 출장 현장 맡김 — 판매 입력에서 가위를 맡아 갈 때 판매와 연결된 복원수리 건을 만든다 (2026-10-27)
 *
 * 돈은 판매(offline_sale_items RS 줄)에만 있다. 이 건은 작업 흐름(꼬리표·검수·내역서·출고) 전용이라
 *   금액은 판매 RS 줄을 그대로 옮겨 표시만 하고, 매출·자루 집계는 sale_id IS NOT NULL 을 빼서 이중 집계를 막는다
 *   (get_hub_stats 165 · use-dashboard-stats · reports/summary · bank/deposit).
 * 결제는 현장에서 끝났다 → 판매가 결제완료면 paid_at 을 바로 채운다(비용안내·입금확인 단계 없음).
 * 가위는 사장님이 들고 오므로 접수확인까지 된 상태(confirmed_at) = 목록의 '입고대기'.
 */
export interface FieldRepairFromSaleInput {
  saleId: string;
  saleNumber: string;
  customerId: string | null;
  name: string;
  phone: string;
  qtyMamoru: number;
  qtyOther: number;
  serviceCost: number;      // 판매 RS 줄 중 수리비 합 (배송비 제외)
  shippingFee: number;      // 판매 RS 줄의 '배송비'
  paid: boolean;            // 판매 결제 상태가 paid 인가
  paymentMethod: string | null;
}

export async function createFieldRepairFromSale(db: Db, input: FieldRepairFromSaleInput) {
  const name = input.name.trim();
  const phone = input.phone.trim();
  const phoneNorm = phone.replace(/\D/g, '');
  const asId = await generateAsId(db);
  const now = new Date().toISOString();

  // 고객 — 판매에서 고른 고객이 있으면 그대로, 없으면 phone 기준 매칭/생성 (SSOT)
  let customerId = input.customerId;
  if (!customerId) {
    ({ customerId } = await matchOrCreateCustomer(db, { phone, name, source: 'manual' }));
  }
  // 단독 송장 발급 대비 — 고객 주소를 복사해 둔다(합포장이면 쓰지 않음)
  let addr: { postcode: string | null; address_road: string | null; address_detail: string | null } | null = null;
  if (customerId) {
    const { data } = await db.from('customers').select('postcode, address_road, address_detail').eq('id', customerId).maybeSingle();
    addr = data || null;
  }

  const { data: repair, error } = await db.from('repairs').insert({
    customer_id: customerId,
    as_id: asId,
    name,
    phone,
    proceed_type: '출장수거',
    postcode: addr?.postcode || null,
    address: addr?.address_road || null,
    address_detail: addr?.address_detail || null,
    qty_mamoru: input.qtyMamoru,
    qty_other: input.qtyOther,
    service_cost: input.serviceCost,
    shipping_fee: input.shippingFee,
    total_amount: input.serviceCost + input.shippingFee,
    status: 'intake',
    received_at: now,
    confirmed_at: now,
    sale_id: input.saleId,
    ...(input.paid ? { paid_at: now, payment_method: input.paymentMethod } : {}),
  }).select().single();
  if (error) throw error;

  await db.from('repair_history').insert({
    repair_id: repair.id,
    to_status: 'intake',
    note: `출장 현장 맡김 — 판매 ${input.saleNumber} 에서 접수${input.paid ? ' (현장 결제 완료)' : ''}`,
  });

  // 맡김 확인 알림톡 — 사장님 본인 등록이라 앱 푸시 없음. 웹훅 미설정(검수 전)이면 sendNotification 이 건너뛴다
  let notified = false;
  try {
    const r = await sendNotification({
      template: 'as_field_received',
      phone: phoneNorm,
      name,
      skipAdminPush: true,
      data: {
        id: asId,
        as_id: asId,
        as_uid: asId,
        qty: String(input.qtyMamoru + input.qtyOther),
        qty_mamoru: String(input.qtyMamoru),
        qty_other: String(input.qtyOther),
        pay_status: input.paid ? '결제 완료' : '판매 건에서 결제 예정',
        // KST 날짜 (UTC slice 금지 — 오전 9시 전 접수가 전날로 찍힌다)
        received_date: formatKoreanDate(new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10)),
      },
    });
    notified = !!r?.success;
  } catch (e) {
    console.error('[repair/field-intake] 알림톡 실패 (접수는 완료):', e);
  }

  return { repair, asId, notified };
}
