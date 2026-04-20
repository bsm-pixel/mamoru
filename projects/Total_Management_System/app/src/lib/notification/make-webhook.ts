/**
 * Make webhook 기반 알림톡 발송 모듈
 * GAS postMake_ 로직과 동일한 payload 형식으로 전송
 *
 * Make 시나리오에서 _meta.func (event) + template 으로 분기 → 솔라피 알림톡
 */

const ENV_WEBHOOK_CONSULTATION = process.env.MAKE_WEBHOOK_URL || '';
const ENV_WEBHOOK_AS_RECEIVED = process.env.MAKE_AS_RECEIVED_WEBHOOK_URL || '';
const ENV_WEBHOOK_REPAIR_STATUS = process.env.MAKE_REPAIR_WEBHOOK_URL || '';
const VERSION = 'tms-2.3';

/**
 * DB 우선 → 환경변수 fallback으로 웹훅 URL 조회
 *
 * 3개 Make 시나리오:
 * 1. consultation  — 상담 접수/확정/취소/리마인더/리뷰
 * 2. as_received   — 복원수리 접수 안내 (별도 시나리오)
 * 3. repair_status — 복원수리 상태변경 (입고확인/입금/출고/취소/만족도)
 */
async function getWebhookUrls(): Promise<{ consultation: string; as_received: string; repair_status: string }> {
  try {
    const { createServiceClient } = require('@/lib/supabase/server');
    const db = createServiceClient();
    const { data: rows } = await db
      .from('system_settings')
      .select('key, value')
      .in('key', ['notifications.webhook_consultation', 'notifications.webhook_as_received', 'notifications.webhook_repair']);

    const map: Record<string, string> = {};
    (rows || []).forEach((r: { key: string; value: string }) => { if (r.value) map[r.key] = String(r.value).replace(/^"|"$/g, ''); });

    return {
      consultation: map['notifications.webhook_consultation'] || ENV_WEBHOOK_CONSULTATION,
      as_received: map['notifications.webhook_as_received'] || ENV_WEBHOOK_AS_RECEIVED,
      repair_status: map['notifications.webhook_repair'] || ENV_WEBHOOK_REPAIR_STATUS,
    };
  } catch {
    return { consultation: ENV_WEBHOOK_CONSULTATION, as_received: ENV_WEBHOOK_AS_RECEIVED, repair_status: ENV_WEBHOOK_REPAIR_STATUS };
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
  | 'review_request'      // 상담 리뷰 요청 → webhook_consultation
  | 'purchase_review_request' // 제품구매 리뷰 요청 → webhook_consultation
  | 'sales_shipped';          // 판매 출고 안내 → webhook_consultation

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
  review_request: 'REVIEW_REQUEST',          // 상담 리뷰 요청 → MAKE_WEBHOOK_URL
  purchase_review_request: 'PURCHASE_REVIEW_REQUEST', // 제품구매 리뷰 요청 → MAKE_WEBHOOK_URL
  sales_shipped: 'SALES_SHIPPED',            // 판매 출고 안내 → MAKE_WEBHOOK_URL
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
  // 설정 기반 on/off 체크
  const enabled = await isNotificationEnabled(payload.template);
  if (!enabled) {
    console.log(`[make-webhook] SKIP template=${payload.template} — 설정에서 비활성`);
    return { success: true }; // 성공으로 처리 (에러 아님)
  }

  // 템플릿에 따라 3분기 웹훅 URL (DB 우선 → 환경변수 fallback)
  const urls = await getWebhookUrls();
  let webhookUrl: string;
  let urlSource: string;

  if (payload.template === 'as_received') {
    // 복원수리 접수 → 별도 Make 시나리오
    webhookUrl = urls.as_received;
    urlSource = 'webhook_as_received';
  } else if (REPAIR_STATUS_TEMPLATES.has(payload.template)) {
    // 복원수리 상태변경 (입고확인/입금/출고/취소/만족도)
    webhookUrl = urls.repair_status;
    urlSource = 'webhook_repair';
  } else {
    // 상담 알림톡 (접수/확정/취소/리마인더/리뷰 등)
    webhookUrl = urls.consultation;
    urlSource = 'webhook_consultation';
  }

  if (!webhookUrl) {
    console.warn(`[make-webhook] SKIP template=${payload.template} — ${urlSource} 미설정 (DB·환경변수 모두 비어있음)`);
    return { success: false, error: `${urlSource} 미설정` };
  }

  // 관리자 푸시 알림 — 웹훅과 무관하게 독립 발송
  const PUSH_CONFIG: Record<string, { title: string; body: string; url: string; settingKey: string }> = {
    confirmed: { title: '새 상담 접수', body: `${payload.name}님 상담 접수`, url: '/consultations', settingKey: 'push.consultation_received' },
    as_received: { title: '새 복원수리 접수', body: `${payload.name}님 복원수리 접수`, url: '/repairs', settingKey: 'push.repair_received' },
    field_request: { title: '새 출장 상담 접수', body: `${payload.name}님 출장 상담 접수`, url: '/consultations', settingKey: 'push.field_request' },
    talk_received: { title: '새 톡상담 접수', body: `${payload.name}님 톡상담 접수`, url: '/consultations', settingKey: 'push.talk_received' },
  };
  const cfg = PUSH_CONFIG[payload.template];
  if (cfg) {
    import('@/lib/firebase/send-push').then(({ sendPushToAll }) => {
      sendPushToAll({
        title: cfg.title,
        body: cfg.body,
        url: cfg.url,
        tag: `mamoru-${payload.template}`,
        settingKey: cfg.settingKey,
      }).catch(() => {});
    }).catch(() => {});
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
