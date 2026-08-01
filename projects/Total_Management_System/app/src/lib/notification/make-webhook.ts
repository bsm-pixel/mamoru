/**
 * Make webhook 기반 알림톡 발송 모듈
 * GAS postMake_ 로직과 동일한 payload 형식으로 전송
 *
 * Make 시나리오에서 _meta.func (event) + template 으로 분기 → 솔라피 알림톡
 */
import { after } from 'next/server';

const ENV_WEBHOOK_CONSULTATION = process.env.MAKE_WEBHOOK_URL || '';
const ENV_WEBHOOK_AS_RECEIVED = process.env.MAKE_AS_RECEIVED_WEBHOOK_URL || '';
const ENV_WEBHOOK_REPAIR_STATUS = process.env.MAKE_REPAIR_WEBHOOK_URL || '';
const ENV_WEBHOOK_EVENT = process.env.MAKE_EVENT_WEBHOOK_URL || '';
const VERSION = 'tms-2.3';

/**
 * DB 우선 → 환경변수 fallback으로 웹훅 URL 조회
 *
 * 4개 Make 시나리오:
 * 1. consultation  — 상담 접수/확정/취소/리마인더/리뷰 (+재고판매)
 * 2. as_received   — 복원수리 접수 안내 (별도 시나리오)
 * 3. repair_status — 복원수리 상태변경 (입고확인/입금/출고/취소/만족도)
 * 4. event         — EVENT 접수확인/입금확인/출고완료 (전용 시나리오, 2026-07-31 분리)
 *                    미설정 시 consultation 로 폴백 → 전환기(웹훅 세팅 전)에도 메시지 유실 방지
 */
async function getWebhookUrls(): Promise<{ consultation: string; as_received: string; repair_status: string; event: string }> {
  try {
    const { createServiceClient } = require('@/lib/supabase/server');
    const db = createServiceClient();
    const { data: rows } = await db
      .from('system_settings')
      .select('key, value')
      .in('key', ['notifications.webhook_consultation', 'notifications.webhook_as_received', 'notifications.webhook_repair', 'notifications.webhook_event']);

    const map: Record<string, string> = {};
    (rows || []).forEach((r: { key: string; value: string }) => { if (r.value) map[r.key] = String(r.value).replace(/^"|"$/g, ''); });

    const consultation = map['notifications.webhook_consultation'] || ENV_WEBHOOK_CONSULTATION;
    return {
      consultation,
      as_received: map['notifications.webhook_as_received'] || ENV_WEBHOOK_AS_RECEIVED,
      repair_status: map['notifications.webhook_repair'] || ENV_WEBHOOK_REPAIR_STATUS,
      // EVENT 전용 미설정 시 consultation 폴백(전환기 안전) — 웹훅 세팅 후엔 완전 분리
      event: map['notifications.webhook_event'] || ENV_WEBHOOK_EVENT || consultation,
    };
  } catch {
    return { consultation: ENV_WEBHOOK_CONSULTATION, as_received: ENV_WEBHOOK_AS_RECEIVED, repair_status: ENV_WEBHOOK_REPAIR_STATUS, event: ENV_WEBHOOK_EVENT || ENV_WEBHOOK_CONSULTATION };
  }
}

// 설정 기반 알림 on/off 체크를 위한 헬퍼
async function isNotificationEnabled(template: string): Promise<boolean> {
  try {
    const { createServiceClient } = require('@/lib/supabase/server');
    const db = createServiceClient();
    // 마스터 스위치
    const { data: master } = await db.from('system_settings').select('value').eq('key', 'notifications.master_enabled').single();
    if (master?.value === 'false' || master?.value === false) return false;
    // 개별 템플릿 매핑
    const templateKeyMap: Record<string, string> = {
      confirmed: 'notifications.consultation_received',
      as_received: 'notifications.repair_received',
      as_cost_notice: 'notifications.repair_cost_notice',
      as_payment_confirmed: 'notifications.repair_payment_confirmed',
      as_shipped: 'notifications.repair_shipped',
      sales_shipped: 'notifications.sales_shipped',
      review_request: 'notifications.review_request',
      event_received: 'notifications.event_received',
      event_payment_notice: 'notifications.event_payment_notice',
      event_payment_confirmed: 'notifications.event_payment_confirmed',
      event_shipped: 'notifications.event_shipped',
      stock_received: 'notifications.stock_received',
      stock_payment_notice: 'notifications.stock_payment_notice',
      stock_payment_confirmed: 'notifications.stock_payment_confirmed',
    };
    const settingKey = templateKeyMap[template];
    if (!settingKey) return true; // 매핑 안 된 템플릿은 항상 발송
    const { data: row } = await db.from('system_settings').select('value').eq('key', settingKey).single();
    if (row?.value === 'false' || row?.value === false) return false;
    return true;
  } catch {
    return true; // DB 오류 시 발송 (안전)
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
  | 'stock_payment_confirmed'; // 재고판매 입금 확인 (→ 판매 자동전환)

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
  /** 🔴 실제 발송이 아니라 설정 토글 OFF 로 '건너뜀'. success=true 지만 고객에겐 안 나감.
   *   발송 시각(shipped_notified_at / review_requested_at)을 찍는 호출부는 이 값을 반드시 확인할 것.
   *   (2026-07-15: 토글 OFF인데 '발송됨'으로 잘못 기록되던 버그 수정) */
  skipped?: boolean;
}> {
  // ── 1) 관리자 앱 푸시 — 고객 행동이면 무조건 발송 ──
  //  ⚠️ 알림톡 설정(on/off)·웹훅 설정과 완전 독립. 반드시 함수 최상단에서 먼저 쏜다.
  //     (알림톡이 꺼져 있어도 사장님 푸시는 항상 울려야 함 — 2026-07-01)
  //  🔔 새 고객 접수/행동 템플릿을 만들면 여기 한 줄만 추가하면 자동으로 울림.
  //     (사장님 자신의 행동=견적발송·출고·입금확인 등은 스팸 방지로 넣지 않음)
  const PUSH_CONFIG: Record<string, { title: string; body: string; url: string; settingKey: string }> = {
    confirmed: { title: '새 상담 접수', body: `${payload.name}님 상담 접수`, url: '/consultations', settingKey: 'push.consultation_received' },
    as_received: { title: '새 복원수리 접수', body: `${payload.name}님 복원수리 접수`, url: '/repairs', settingKey: 'push.repair_received' },
    // 출장 신규: submit/route.ts 는 template='request' 로 호출 (솔라피 템플릿명과 일치)
    request: { title: '새 출장 상담 접수', body: `${payload.name}님 출장 상담 접수`, url: '/consultations', settingKey: 'push.field_request' },
    field_request: { title: '새 출장 상담 접수', body: `${payload.name}님 출장 상담 접수`, url: '/consultations', settingKey: 'push.field_request' },
    talk_received: { title: '새 톡상담 접수', body: `${payload.name}님 톡상담 접수`, url: '/consultations', settingKey: 'push.talk_received' },
    // 취소: 고객이 page_change_request 에서 취소 → public/cancel/route.ts
    field_cancelled: { title: '⚠️ 출장 예약 취소', body: `${payload.name}님 출장 예약 취소`, url: '/consultations', settingKey: 'push.field_cancelled' },
    cancelled: { title: '⚠️ 상담 예약 취소', body: `${payload.name}님 상담 예약 취소`, url: '/consultations', settingKey: 'push.consultation_cancelled' },
    // 이벤트 접수(고객) — 2026-07-01 추가
    event_received: { title: '새 이벤트 접수', body: `${payload.name}님 이벤트 접수`, url: '/events', settingKey: 'push.event_received' },
    // 재고판매 접수(고객) — 2026-07-21 추가
    stock_received: { title: '새 재고판매 접수', body: `${payload.name}님 재고판매 주문`, url: '/stock-sale', settingKey: 'push.stock_received' },
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

  // ── 2) 고객 알림톡 on/off (푸시엔 영향 없음 — 위에서 이미 발송) ──
  const enabled = await isNotificationEnabled(payload.template);
  if (!enabled) {
    console.log(`[make-webhook] SKIP 알림톡 template=${payload.template} — 설정에서 비활성 (관리자 푸시는 발송됨)`);
    return { success: true, skipped: true }; // 에러는 아니지만 '실제 발송 X' → skipped 로 명시 (발송시각 오기록 방지)
  }

  // 템플릿에 따라 3분기 웹훅 URL (DB 우선 → 환경변수 fallback)
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
