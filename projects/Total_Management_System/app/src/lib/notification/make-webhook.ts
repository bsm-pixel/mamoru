/**
 * Make webhook 기반 알림톡 발송 모듈
 * GAS postMake_ 로직과 동일한 payload 형식으로 전송
 *
 * Make 시나리오에서 _meta.func (event) + template 으로 분기 → 솔라피 알림톡
 */

const MAKE_WEBHOOK_URL = process.env.MAKE_WEBHOOK_URL || '';           // 상담 알림톡
const MAKE_REPAIR_WEBHOOK_URL = process.env.MAKE_REPAIR_WEBHOOK_URL || ''; // 복원수리 상태변경
const VERSION = 'tms-2.2';

/** 복원수리 상태변경 전용 템플릿 (별도 Make 시나리오) */
const REPAIR_STATUS_TEMPLATES = new Set<NotifyTemplate>([
  'as_cost_notice',
  'as_payment_confirmed',
  'as_shipped',
  'as_satisfaction',
]);

export type NotifyTemplate =
  | 'confirmed'           // 매장방문 확정
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
  | 'as_received'         // 복원수리 접수 안내
  | 'as_cost_notice'      // 비용 안내
  | 'as_payment_confirmed' // 입금 확인
  | 'as_shipped'          // 출고 안내
  | 'as_satisfaction';    // 만족도/리뷰 요청

/** GAS postMake_ event명 매핑 */
const TEMPLATE_EVENT_MAP: Record<NotifyTemplate, string> = {
  confirmed: 'CONFIRMED',
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
  as_satisfaction: 'AS_SATISFACTION',
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
  // 템플릿에 따라 웹훅 URL 분기
  const isRepairStatus = REPAIR_STATUS_TEMPLATES.has(payload.template);
  const webhookUrl = isRepairStatus ? MAKE_REPAIR_WEBHOOK_URL : MAKE_WEBHOOK_URL;
  const envName = isRepairStatus ? 'MAKE_REPAIR_WEBHOOK_URL' : 'MAKE_WEBHOOK_URL';

  if (!webhookUrl) {
    return { success: false, error: `${envName} 환경변수 미설정` };
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
        return { success: true };
      }

      // 429, 5xx → 재시도
      const shouldRetry = res.status === 429 || (res.status >= 500 && res.status < 600);
      if (!shouldRetry) {
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
