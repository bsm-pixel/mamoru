/**
 * 사장님 알림 메일 — 관리자 푸시의 "예비 채널" (2026-09-29)
 *
 * 푸시가 휴대폰에 도착하지 않았을 때 bsm@mamoru.kr 로 메일 → Gmail 앱(네이티브 앱) 알림.
 * 브라우저·TMS 로그인 상태와 무관하게 울린다. (처음엔 솔라피 문자로 설계 → 사장님 결정으로 메일, 비용 0)
 * Make 를 거치지 않는다 — Make 는 순단 1회로 시나리오가 꺼진 전례(2026-09-17)가 있어 안전망이 기대면 안 된다.
 *
 * 받는 주소: OWNER_ALERT_EMAIL(선택) → 구글 연결 계정 → bsm@mamoru.kr
 */

import { sendMail } from '@/lib/google/gmail-client';
import { getConnectionStatus } from '@/lib/google/oauth';

const GMAIL_SEND_SCOPE = 'https://www.googleapis.com/auth/gmail.send';
const DEFAULT_TO = 'bsm@mamoru.kr';
const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';

/** 메일 안전망 상태 — 설정 화면 경고용 */
export async function getOwnerAlertStatus(): Promise<{ ready: boolean; problem?: string }> {
  try {
    const s = await getConnectionStatus();
    if (!s.connected) return { ready: false, problem: '구글 계정이 연결돼 있지 않습니다 (설정 → 구글 캘린더 연결).' };
    // granted_scopes 가 기록된 연결이면 gmail.send 허용 여부까지 확인
    if (s.granted_scopes && !s.granted_scopes.includes(GMAIL_SEND_SCOPE)) {
      return { ready: false, problem: '구글 연결에 메일 보내기 권한이 없습니다 — 설정 → 구글 캘린더에서 재연결 1회.' };
    }
    return { ready: true };
  } catch {
    return { ready: false, problem: '구글 연결 상태를 확인하지 못했습니다.' };
  }
}

/** 사장님께 알림 메일 1통. 실패해도 throw 하지 않고 결과 문자열을 돌려준다(호출처가 기록) */
export async function sendOwnerAlert(subject: string, body: string, path?: string): Promise<{ ok: boolean; result: string }> {
  let to = (process.env.OWNER_ALERT_EMAIL || '').trim();
  if (!to) {
    try { to = (await getConnectionStatus()).email || DEFAULT_TO; } catch { to = DEFAULT_TO; }
  }
  const text = path ? `${body}\n\nTMS 열기: ${BASE_URL}${path}` : body;
  const r = await sendMail({ to, subject, text });
  if (r.ok) return { ok: true, result: 'mailed' };
  const kind = r.notConnected ? 'not-connected' : r.apiDisabled ? 'gmail-api-disabled' : r.needsReauth ? 'needs-reauth' : 'error';
  console.error('[owner-alert] 메일 실패:', kind, r.error);
  return { ok: false, result: `${kind}: ${(r.error || '').slice(0, 200)}` };
}
