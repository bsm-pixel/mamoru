/**
 * 사장님 휴대폰 문자 — 관리자 푸시의 "예비 채널" (2026-09-29)
 *
 * 왜 Make 를 거치지 않고 솔라피를 직접 부르나:
 *   이 문자는 "푸시가 실패했을 때의 마지막 안전망"이다. Make 는 솔라피 연결 순단 1회로
 *   시나리오 전체가 꺼지는 전례가 있어(2026-09-17) 안전망이 안전망에 기대면 안 된다.
 *   → 솔라피 REST API 직접 호출 (HMAC-SHA256 인증). 템플릿 검수 없는 일반 문자라 즉시 발송 가능.
 *
 * 필요한 환경변수 (Vercel):
 *   SOLAPI_API_KEY / SOLAPI_API_SECRET — 솔라피 콘솔 > 개발/연동 > API Key
 *   SOLAPI_SENDER                      — 솔라피에 등록된 발신번호
 *   OWNER_ALERT_PHONE                  — 문자를 받을 사장님 휴대폰 번호
 */

import crypto from 'crypto';

const digits = (v: string | undefined) => (v || '').replace(/[^0-9]/g, '');

function getConfig() {
  const apiKey = (process.env.SOLAPI_API_KEY || '').trim();
  const apiSecret = (process.env.SOLAPI_API_SECRET || '').trim();
  const from = digits(process.env.SOLAPI_SENDER);
  const to = digits(process.env.OWNER_ALERT_PHONE);
  return { apiKey, apiSecret, from, to };
}

/** 문자 예비채널이 설정돼 있는가 — 설정 화면 경고 표시용 */
export function isOwnerSmsReady(): boolean {
  const c = getConfig();
  return !!(c.apiKey && c.apiSecret && c.from && c.to);
}

/**
 * 사장님에게 문자 1통. 길이에 따라 솔라피가 SMS/LMS 를 자동 판별한다.
 * 실패해도 throw 하지 않는다 — 결과 문자열을 호출처가 기록한다.
 */
export async function sendOwnerSms(text: string): Promise<{ ok: boolean; result: string }> {
  const { apiKey, apiSecret, from, to } = getConfig();
  if (!apiKey || !apiSecret || !from || !to) {
    console.warn('[owner-sms] 환경변수 미설정 — 문자 예비발송 불가');
    return { ok: false, result: 'env-missing' };
  }

  const date = new Date().toISOString();
  const salt = crypto.randomBytes(16).toString('hex');
  const signature = crypto.createHmac('sha256', apiSecret).update(date + salt).digest('hex');

  try {
    const res = await fetch('https://api.solapi.com/messages/v4/send', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `HMAC-SHA256 apiKey=${apiKey}, date=${date}, salt=${salt}, signature=${signature}`,
      },
      body: JSON.stringify({ message: { to, from, text } }),
    });
    const body = await res.text();
    if (!res.ok) {
      console.error('[owner-sms] 발송 실패:', res.status, body.slice(0, 300));
      return { ok: false, result: `http-${res.status}: ${body.slice(0, 200)}` };
    }
    return { ok: true, result: 'sent' };
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    console.error('[owner-sms] 발송 오류:', msg);
    return { ok: false, result: `error: ${msg.slice(0, 200)}` };
  }
}
