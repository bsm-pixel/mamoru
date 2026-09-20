/**
 * Make webhook 기반 알림톡 발송 모듈
 * GAS postMake_ 로직과 동일한 payload 형식으로 전송
 *
 * Make 시나리오에서 _meta.func (event) + template 으로 분기 → 솔라피 알림톡
 *
 * 📌 발송 원칙(2026-09-14 사장님 확정): **항상 발송** — 설정 on/off 토글 없음.
 *   고객 행동·상태 변화마다 빠짐없이 안내한다. 같은 내용이 두 번 가는 경우만 호출부의 중복 방지 규칙으로 막는다.
 *   비상 시 특정 흐름을 멈추려면 설정의 해당 Make 웹훅 URL 을 비우거나 Make 시나리오를 끈다.
 */
import { after } from 'next/server';

const ENV_WEBHOOK_CONSULTATION = process.env.MAKE_WEBHOOK_URL || '';
const ENV_WEBHOOK_AS_RECEIVED = process.env.MAKE_AS_RECEIVED_WEBHOOK_URL || '';
const ENV_WEBHOOK_REPAIR_STATUS = process.env.MAKE_REPAIR_WEBHOOK_URL || '';
const ENV_WEBHOOK_EVENT = process.env.MAKE_EVENT_WEBHOOK_URL || '';
const ENV_WEBHOOK_IMWEB = process.env.MAKE_IMWEB_WEBHOOK_URL || '';
const VERSION = 'tms-2.3';

/**
 * DB 우선 → 환경변수 fallback으로 웹훅 URL 조회
 *
 * 5개 Make 시나리오:
 * 1. consultation  — 상담 접수/확정/취소/리마인더/리뷰 (+재고판매)
 * 2. as_received   — 복원수리 접수 안내 (별도 시나리오)
 * 3. repair_status — 복원수리 상태변경 (입고확인/입금/출고/취소/만족도)
 * 4. event         — EVENT 접수확인/입금확인/출고완료 (전용 시나리오, 2026-07-31 분리)
 *                    미설정 시 consultation 로 폴백 → 전환기(웹훅 세팅 전)에도 메시지 유실 방지
 * 5. imweb         — 아임웹 주문 취소·반품 안내 (2026-09-14). **폴백 없음**: 미설정이면 발송하지 않음
 *                    (Make 분기·솔라피 승인 전 오발송 방지 — URL 입력이 곧 가동 시점)
 */
async function getWebhookUrls(): Promise<{ consultation: string; as_received: string; repair_status: string; event: string; imweb: string; sales: string; returns: string }> {
  try {
    const { createServiceClient } = await import('@/lib/supabase/server');
    const db = createServiceClient();
    const { data: rows } = await db
      .from('system_settings')
      .select('key, value')
      .in('key', ['notifications.webhook_consultation', 'notifications.webhook_as_received', 'notifications.webhook_repair', 'notifications.webhook_event', 'notifications.webhook_imweb', 'notifications.webhook_sales', 'notifications.webhook_returns']);

    const map: Record<string, string> = {};
    (rows || []).forEach((r: { key: string; value: string }) => { if (r.value) map[r.key] = String(r.value).replace(/^"|"$/g, ''); });

    const consultation = map['notifications.webhook_consultation'] || ENV_WEBHOOK_CONSULTATION;
    return {
      consultation,
      as_received: map['notifications.webhook_as_received'] || ENV_WEBHOOK_AS_RECEIVED,
      repair_status: map['notifications.webhook_repair'] || ENV_WEBHOOK_REPAIR_STATUS,
      // EVENT 전용 미설정 시 consultation 폴백(전환기 안전) — 웹훅 세팅 후엔 완전 분리
      event: map['notifications.webhook_event'] || ENV_WEBHOOK_EVENT || consultation,
      // 아임웹 주문 안내는 폴백 없음 — 미설정이면 미발송(가동 전 오발송 방지)
      imweb: map['notifications.webhook_imweb'] || ENV_WEBHOOK_IMWEB,
      // 2026-09-20 신설 — 미설정이면 consultation 으로(= 지금과 동일). 새 시나리오 URL 넣으면 전환
      sales: map['notifications.webhook_sales'] || consultation,
      returns: map['notifications.webhook_returns'] || consultation,
    };
  } catch {
    return { consultation: ENV_WEBHOOK_CONSULTATION, as_received: ENV_WEBHOOK_AS_RECEIVED, repair_status: ENV_WEBHOOK_REPAIR_STATUS, event: ENV_WEBHOOK_EVENT || ENV_WEBHOOK_CONSULTATION, imweb: ENV_WEBHOOK_IMWEB, sales: ENV_WEBHOOK_CONSULTATION, returns: ENV_WEBHOOK_CONSULTATION };
  }
}

/** 복원수리 상태변경 전용 템플릿 (별도 Make 시나리오 → MAKE_REPAIR_WEBHOOK_URL) */
const REPAIR_STATUS_TEMPLATES = new Set<NotifyTemplate>([
  'as_cost_notice',
  'as_payment_confirmed',
  'as_shipped',
  'as_cancelled',
  'as_review_request',  // 복원수리 리뷰 요청
  // Phase 4: 직접방문 (booked 는 as_received URL 로 별도 라우팅)
  'as_visit_remind_24h',
  'as_visit_remind_2h',
  'as_visit_rescheduled',
  'as_visit_cancelled',
]);

/** 아임웹 주문 취소·반품 안내 (전용 Make 시나리오 → webhook_imweb, 폴백 없음) — 2026-09-14 */
const IMWEB_ORDER_TEMPLATES = new Set<NotifyTemplate>([
  'imweb_cancel_requested',
  'imweb_cancel_completed',
  'imweb_return_requested',
  'imweb_return_approved',
  'imweb_return_completed',
]);

/* ── Make 시나리오 분리 (2026-09-20) ────────────────────────────────────────
   전엔 `상담접수_제작중` 한 시나리오에 상담 14 + 후기 3 + 출고 + 반품 2 + 지연 = 20종 넘게 몰려 있었다.
   2026-09-17 사고: 솔라피 연결 순단 1회로 그 시나리오가 꺼지면서 **저 20종이 전부 멈췄다**.
   → 고객 흐름 단위로 쪼개 **장애를 격리**한다.

   🔒 전환 안전장치: 새 키가 비어 있으면 **기존대로 consultation 으로 간다**.
      그래서 이 배포만으로는 동작이 하나도 안 바뀌고, 새 시나리오 URL 을 넣는 순간 전환된다.
      문제가 생기면 URL 만 비우면 즉시 롤백(무중단).

   ※ 복원수리 접수(as_received)와 상태변경을 한 시나리오로 합치고 싶으면
      **설정에서 webhook_as_received 에 webhook_repair 와 같은 URL 을 넣으면 된다** (코드 변경 불필요). */

/** 판매·배송 — 아임웹 주문·오프라인 판매 공통 (출고·구매후기) */
const SALES_TEMPLATES = new Set<NotifyTemplate>([
  'sales_shipped',
  'purchase_review_request',
]);

/** 반품·교환 — TMS /returns (오프라인 판매분). 아임웹 주문 클레임은 IMWEB_ORDER_TEMPLATES */
const RETURN_TEMPLATES = new Set<NotifyTemplate>([
  'return_received',
  'return_inbound',
]);

/** EVENT 전용 템플릿 (전용 Make 시나리오 → webhook_event) — 2026-07-31 consultation 에서 분리 */
const EVENT_TEMPLATES = new Set<NotifyTemplate>([
  'event_received',
  'event_payment_notice',
  'event_payment_confirmed',
  'event_shipped',
]);

export type NotifyTemplate =
  | 'confirmed'           // 매장방문 확정
  | 'request'             // 출장요청 접수 안내
  | 'cancelled'           // 매장 취소 안내
  | 'suggest'             // 출장 시간 제안 (SUGGESTED_TIMES)
  | 'rescheduled'         // 매장 일정 변경
  | 'field_confirmed'     // 출장 확정
  | 'field_cancelled'     // 출장 취소 안내
  | 'field_rescheduled'   // 출장 일정 변경
  | 'field_remind_24h'    // 출장 24h 리마인드
  | 'field_remind_2h'     // 출장 2h 리마인드
  | 'remind24'            // 매장방문 24h 리마인드 (Make 필터: template equal to remind24)
  | 'remind2'             // 매장방문 2h 리마인드 (Make 필터: template equal to remind2)
  | 'field_delayed'       // 출장 지연 안내
  | 'talk_received'       // 톡상담 접수 안내
  | 'talk_ready'          // 톡상담 시작 안내
  // Phase 7: 복원수리 알림톡
  | 'as_received'         // 복원수리 접수 안내 → webhook_as_received (별도 시나리오)
  | 'as_cost_notice'      // 비용 안내 → webhook_repair
  | 'as_payment_confirmed' // 입금 확인 → webhook_repair
  | 'as_shipped'          // 출고 안내 → webhook_repair
  | 'as_cancelled'        // 복원수리 취소 안내 → webhook_repair
  | 'as_review_request'   // 복원수리 만족도 → webhook_repair
  // Phase 4: 직접방문(당일수리) — booked=as_received URL, 나머지=repair_status URL
  | 'as_visit_booked'     // 직접방문 접수완료 → webhook_as_received
  | 'as_visit_remind_24h' // 직접방문 D-1 리마인드 → webhook_repair
  | 'as_visit_remind_2h'  // 직접방문 당일 2h 리마인드 → webhook_repair
  | 'as_visit_rescheduled' // 직접방문 예약 변경 안내 → webhook_repair
  | 'as_visit_cancelled'  // 직접방문 예약 취소 완료 → webhook_repair
  | 'review_request'      // 상담 리뷰 요청 → webhook_consultation
  | 'purchase_review_request' // 제품구매 리뷰 요청 → webhook_consultation
  | 'sales_shipped'           // 판매 출고 안내 → webhook_consultation
  // EVENT(고객 접수) — webhook_event 전용 시나리오 (2026-07-31 분리)
  | 'event_received'          // EVENT 접수 확인 + 비용안내 (자동)
  | 'event_payment_notice'    // EVENT 입금 안내 (총액+계좌, 사장님 재고확인 후)
  | 'event_payment_confirmed' // EVENT 입금 확인 (→ 판매 자동전환)
  | 'event_shipped'           // EVENT 출고완료 (판매전환분 출고 시 자동, sales_shipped 대체)
  // 재고판매(LS) — webhook_consultation 시나리오 사용
  | 'stock_received'          // 재고판매 접수 확인 + 입금 안내(계좌+금액) (자동)
  | 'stock_payment_notice'    // 재고판매 입금 안내 재발송 (어드민)
  | 'stock_payment_confirmed'  // 재고판매 입금 확인 (→ 판매 자동전환)
  // 반품·교환수거 (2026-08-25) — webhook_consultation 폴백
  | 'return_received'          // 반품수거 접수 (교환/반품 시 자동, 사장님 푸시)
  | 'return_inbound'           // 반품 입고완료 (사장님 처리 시 고객 알림)
  // 아임웹 주문 취소·반품 (2026-09-14) — webhook_imweb 전용 · 솔라피 IW- 템플릿
  | 'imweb_cancel_requested'   // IW-취소접수 (고객 요청만 — 관리자 직접 취소는 생략)
  | 'imweb_cancel_completed'   // IW-취소완료 (환불 안내 겸)
  | 'imweb_return_requested'   // IW-반품접수 (고객 요청만)
  | 'imweb_return_approved'    // IW-반품승인 (반품 수거중 = 승인)
  | 'imweb_return_completed';  // IW-반품완료 (환불 안내 겸)

/** GAS postMake_ event명 매핑 */
const TEMPLATE_EVENT_MAP: Record<NotifyTemplate, string> = {
  confirmed: 'CONFIRMED',
  request: 'CONSULT_REQUEST',
  cancelled: 'CANCELLED',
  suggest: 'SUGGESTED_TIMES',
  rescheduled: 'RESCHEDULED',
  field_confirmed: 'FIELD_CONFIRMED',
  field_cancelled: 'FIELD_CANCELLED',
  field_rescheduled: 'FIELD_RESCHEDULED',
  field_remind_24h: 'FIELD_REMIND_24H',
  field_remind_2h: 'FIELD_REMIND_2H',
  remind24: 'REMIND_24H',
  remind2: 'REMIND_2H',
  field_delayed: 'FIELD_DELAYED',
  talk_received: 'TALK_RECEIVED',
  talk_ready: 'TALK_READY',
  // Phase 7: 복원수리
  as_received: 'AS_RECEIVED',
  as_cost_notice: 'AS_COST_NOTICE',
  as_payment_confirmed: 'AS_PAYMENT_CONFIRMED',
  as_shipped: 'AS_SHIPPED',
  as_cancelled: 'AS_CANCELLED',
  as_review_request: 'AS_REVIEW_REQUEST',   // 복원수리 리뷰 요청 → MAKE_REPAIR_WEBHOOK_URL
  // Phase 4: 직접방문(당일수리)
  as_visit_booked: 'AS_VISIT_BOOKED',
  as_visit_remind_24h: 'AS_VISIT_REMIND_24H',
  as_visit_remind_2h: 'AS_VISIT_REMIND_2H',
  as_visit_rescheduled: 'AS_VISIT_RESCHEDULED',
  as_visit_cancelled: 'AS_VISIT_CANCELLED',
  review_request: 'REVIEW_REQUEST',          // 상담 리뷰 요청 → MAKE_WEBHOOK_URL
  purchase_review_request: 'PURCHASE_REVIEW_REQUEST', // 제품구매 리뷰 요청 → MAKE_WEBHOOK_URL
  sales_shipped: 'SALES_SHIPPED',            // 판매 출고 안내 → MAKE_WEBHOOK_URL
  // EVENT
  event_received: 'EVENT_RECEIVED',
  event_payment_notice: 'EVENT_PAYMENT_NOTICE',
  event_payment_confirmed: 'EVENT_PAYMENT_CONFIRMED',
  event_shipped: 'EVENT_SHIPPED',
  // 재고판매(LS)
  stock_received: 'STOCK_RECEIVED',
  stock_payment_notice: 'STOCK_PAYMENT_NOTICE',
  stock_payment_confirmed: 'STOCK_PAYMENT_CONFIRMED',
  // 반품·교환수거
  return_received: 'RETURN_RECEIVED',
  return_inbound: 'RETURN_INBOUND',
  // 아임웹 주문 취소·반품
  imweb_cancel_requested: 'IMWEB_CANCEL_REQUESTED',
  imweb_cancel_completed: 'IMWEB_CANCEL_COMPLETED',
  imweb_return_requested: 'IMWEB_RETURN_REQUESTED',
  imweb_return_approved: 'IMWEB_RETURN_APPROVED',
  imweb_return_completed: 'IMWEB_RETURN_COMPLETED',
};

interface NotifyPayload {
  template: NotifyTemplate;
  phone: string;
  name: string;
  /** 추가 데이터 — GAS payload와 동일 키 사용 (date, time, address 등) */
  data?: Record<string, string>;
}

/** 간단한 UUID 생성 */
function uuid(): string {
  return crypto.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

export async function sendNotification(payload: NotifyPayload): Promise<{
  success: boolean;
  error?: string;
}> {
  // ── 1) 관리자 앱 푸시 — 고객 행동이면 무조건 발송 ──
  //  ⚠️ 웹훅 설정과 완전 독립. 반드시 함수 최상단에서 먼저 쏜다. (웹훅 미설정이어도 사장님 푸시는 항상 울림)
  //  🔔 새 고객 접수/행동 템플릿을 만들면 여기 한 줄만 추가하면 자동으로 울림.
  //     (사장님 자신의 행동=견적발송·출고·입금확인 등은 스팸 방지로 넣지 않음)
  const PUSH_CONFIG: Record<string, { title: string; body: string; url: string }> = {
    confirmed: { title: '새 상담 접수', body: `${payload.name}님 상담 접수`, url: '/consultations' },
    as_received: { title: '새 복원수리 접수', body: `${payload.name}님 복원수리 접수`, url: '/repairs' },
    // 직접방문(매장방문) 접수
    as_visit_booked: { title: '새 매장방문 수리 접수', body: `${payload.name}님 매장방문 수리 접수`, url: '/repairs' },
    // 🔴 2026-09-15 누락 복구 — 고객이 직접방문 예약을 스스로 취소/변경한 건.
    //    슬롯이 비거나 시간이 바뀌므로 사장님이 반드시 알아야 하는데 푸시가 없었다.
    as_visit_cancelled: { title: '⚠️ 매장방문 수리 취소', body: `${payload.name}님 매장방문 예약 취소`, url: '/repairs' },
    as_visit_rescheduled: { title: '매장방문 일정 변경', body: `${payload.name}님 방문 일정 변경`, url: '/repairs' },
    // 출장 신규: submit/route.ts 는 template='request' 로 호출 (솔라피 템플릿명과 일치)
    request: { title: '새 출장 상담 접수', body: `${payload.name}님 출장 상담 접수`, url: '/consultations' },
    field_request: { title: '새 출장 상담 접수', body: `${payload.name}님 출장 상담 접수`, url: '/consultations' },
    talk_received: { title: '새 톡상담 접수', body: `${payload.name}님 톡상담 접수`, url: '/consultations' },
    // 취소: 고객이 page_change_request 에서 취소 → public/cancel/route.ts
    field_cancelled: { title: '⚠️ 출장 예약 취소', body: `${payload.name}님 출장 예약 취소`, url: '/consultations' },
    cancelled: { title: '⚠️ 상담 예약 취소', body: `${payload.name}님 상담 예약 취소`, url: '/consultations' },
    // 이벤트 접수(고객) — 2026-07-01 추가
    event_received: { title: '새 이벤트 접수', body: `${payload.name}님 이벤트 접수`, url: '/events' },
    // 재고판매 접수(고객) — 2026-07-21 추가
    stock_received: { title: '새 재고판매 접수', body: `${payload.name}님 재고판매 주문`, url: '/stock-sale' },
    // 반품·교환수거 접수 — 2026-08-25 추가 (사장님 푸시)
    return_received: { title: '새 반품·교환수거 접수', body: `${payload.name}님 반품수거 접수`, url: '/returns' },
    // 아임웹 주문 취소·반품 요청(고객) — 사장님이 아임웹에서 승인/거절 처리해야 함 (2026-09-14)
    imweb_cancel_requested: { title: '아임웹 주문 취소 요청', body: `${payload.name}님 취소 요청 · 아임웹에서 승인/거절 처리`, url: '/orders' },
    imweb_return_requested: { title: '아임웹 주문 반품 요청', body: `${payload.name}님 반품 요청 · 아임웹에서 승인/거절 처리`, url: '/orders' },
  };
  const pushCfg = PUSH_CONFIG[payload.template];
  if (pushCfg) {
    // tag에 건별 고유 ID 포함 — 레코드 삭제 시 SW에서 해당 알림만 정확히 회수 가능
    const uniqId = payload.data?.as_id || payload.data?.id || '';
    const pushTag = uniqId ? `mamoru-${payload.template}-${uniqId}` : `mamoru-${payload.template}`;
    // 🔴 완주 보장 — after()로 넘겨야 Vercel이 응답 후 함수를 종료시켜도 푸시가 끝까지 발송된다.
    //    (기존 fire-and-forget `import().then()` 은 접수 3종 모바일 알림이 오락가락하던 근본원인 — 2026-08-01)
    const deliverPush = async () => {
      const { sendPushToAll } = await import('@/lib/firebase/send-push');
      await sendPushToAll({ title: pushCfg.title, body: pushCfg.body, url: pushCfg.url, tag: pushTag });
    };
    try {
      after(deliverPush);                    // 요청 컨텍스트: 응답 후 플랫폼이 함수를 살려 완주
    } catch {
      await deliverPush().catch(() => {});   // 요청 밖(크론 등): 인라인으로 완주
    }
  }

  // ── 2) 고객 알림톡 — 항상 발송 (on/off 토글 없음, 2026-09-14) ──
  // 템플릿에 따라 웹훅 URL 분기 (DB 우선 → 환경변수 fallback)
  const urls = await getWebhookUrls();
  let webhookUrl: string;
  let urlSource: string;

  if (payload.template === 'as_received' || payload.template === 'as_visit_booked') {
    // 복원수리 접수(택배·방문) → 별도 Make 시나리오 (접수 알림)
    webhookUrl = urls.as_received;
    urlSource = 'webhook_as_received';
  } else if (REPAIR_STATUS_TEMPLATES.has(payload.template)) {
    // 복원수리 상태변경 (입고확인/입금/출고/취소/만족도)
    webhookUrl = urls.repair_status;
    urlSource = 'webhook_repair';
  } else if (EVENT_TEMPLATES.has(payload.template)) {
    // EVENT 접수확인/입금확인/출고완료 → 전용 시나리오 (미설정 시 consultation 폴백)
    webhookUrl = urls.event;
    urlSource = 'webhook_event';
  } else if (SALES_TEMPLATES.has(payload.template)) {
    webhookUrl = urls.sales;
    urlSource = 'webhook_sales';
  } else if (RETURN_TEMPLATES.has(payload.template)) {
    webhookUrl = urls.returns;
    urlSource = 'webhook_returns';
  } else if (IMWEB_ORDER_TEMPLATES.has(payload.template)) {
    // 아임웹 주문 취소·반품 → 전용 시나리오 (폴백 없음)
    webhookUrl = urls.imweb;
    urlSource = 'webhook_imweb';
  } else {
    // 상담 알림톡 (접수/확정/취소/리마인더/리뷰 등)
    webhookUrl = urls.consultation;
    urlSource = 'webhook_consultation';
  }

  // 고객 알림톡 웹훅 미설정이면 여기서 종료 (관리자 푸시는 함수 시작부에서 이미 발송됨)
  if (!webhookUrl) {
    console.warn(`[make-webhook] SKIP 알림톡 template=${payload.template} — ${urlSource} 미설정 (관리자 푸시는 발송됨)`);
    return { success: false, error: `${urlSource} 미설정` };
  }

  const event = TEMPLATE_EVENT_MAP[payload.template] || payload.template.toUpperCase();
  const corrId = uuid();
  const uid = payload.data?.id || uuid();
  const idemKey = `${uid}:${payload.template}:${payload.data?.date || 'na'}T${payload.data?.time || 'na'}`;

  // GAS postMake_와 동일한 구조
  const body = {
    _meta: {
      ts: new Date().toISOString(),
      version: VERSION,
      func: event,
      trigger: 'tms',
    },
    topic: 'alrimtalk',
    template: payload.template,
    event,
    name: payload.name,
    phone: payload.phone.replace(/\D/g, ''),
    channel: 'kakao',
    sms_fallback: false,
    ...payload.data,
  };

  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    'Accept': 'application/json',
    'X-Correlation-Id': corrId,
    'X-Idempotency-Key': idemKey,
  };

  // 3회 재시도 (GAS와 동일)
  let lastStatus = 0;
  let lastBody = '';

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const res = await fetch(webhookUrl, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(5000),
      });

      lastStatus = res.status;
      lastBody = await res.text();

      if (res.ok) {
        console.log(`[make-webhook] OK template=${payload.template} phone=${payload.phone.slice(-4)} status=${res.status}`);
        return { success: true };
      }

      // 429, 5xx → 재시도
      const shouldRetry = res.status === 429 || (res.status >= 500 && res.status < 600);
      if (!shouldRetry) {
        console.error(`[make-webhook] FAIL template=${payload.template} status=${res.status} body=${lastBody.slice(0, 200)}`);
        return { success: false, error: `Make HTTP ${res.status}: ${lastBody}` };
      }
    } catch (err) {
      lastBody = String(err);
      if (attempt >= 3) {
        return { success: false, error: `${lastStatus ? `HTTP ${lastStatus} ` : ''}${lastBody}` };
      }
    }

    // 백오프 대기 (500ms, 1000ms, 2000ms + jitter)
    await new Promise((r) => setTimeout(r, 500 * Math.pow(2, attempt - 1) + Math.random() * 250));
  }

  return { success: false, error: `재시도 초과 HTTP ${lastStatus}: ${lastBody}` };
}
