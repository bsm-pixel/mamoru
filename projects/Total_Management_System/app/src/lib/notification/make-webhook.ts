/**
 * Make webhook 기반 알림톡 발송 모듈
 * GAS postMake_ 로직을 TMS로 이식
 */

const MAKE_WEBHOOK_URL = process.env.MAKE_WEBHOOK_URL || '';

export type NotifyTemplate =
  | 'confirmed'           // 매장방문 확정
  | 'cancelled'           // 취소 안내
  | 'suggest'             // 출장 시간 제안
  | 'rescheduled'         // 일정 변경
  | 'field_confirmed';    // 출장 확정

interface NotifyPayload {
  template: NotifyTemplate;
  phone: string;
  name: string;
  /** 추가 데이터 (날짜, 시간, 주소 등) */
  data?: Record<string, string>;
}

export async function sendNotification(payload: NotifyPayload): Promise<{
  success: boolean;
  error?: string;
}> {
  if (!MAKE_WEBHOOK_URL) {
    return { success: false, error: 'MAKE_WEBHOOK_URL 환경변수 미설정' };
  }

  try {
    const body = {
      template: payload.template,
      phone: payload.phone.replace(/\D/g, ''),
      name: payload.name,
      ...payload.data,
    };

    const res = await fetch(MAKE_WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const text = await res.text();
      return { success: false, error: `Make 응답 ${res.status}: ${text}` };
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: String(err) };
  }
}
