/**
 * Gmail 발송 (2026-09-29) — 관리자 푸시 미수신 시 사장님께 알림 메일
 *
 * 기존 구글 연결(캘린더·할 일과 같은 토큰)을 그대로 쓴다. `gmail.send` 스코프가 필요 → 설정에서 재연결 1회.
 * 사장님 계정 → 사장님 계정("나에게 보내기")이어도 Gmail 앱 알림이 울리는 것을 사장님이 직접 확인함(2026-09-29).
 * 실패는 throw 하지 않고 결과로 반환한다 — 호출처가 기록한다.
 */

import { google } from 'googleapis';
import { getAuthorizedClient } from './oauth';

export interface MailResult {
  ok: boolean;
  error?: string;
  notConnected?: boolean;
  needsReauth?: boolean;   // 토큰에 gmail.send 권한 없음 → 설정에서 재연결 1회
  apiDisabled?: boolean;   // 구글 클라우드에서 Gmail API 꺼짐 → 콘솔에서 '사용 설정' (재연결로 안 풀림)
}

function isApiDisabled(msg: string): boolean {
  return /has not been used in project|is disabled|SERVICE_DISABLED|accessNotConfigured/i.test(msg);
}
function isScopeError(msg: string): boolean {
  if (isApiDisabled(msg)) return false;
  return /insufficient|ACCESS_TOKEN_SCOPE|invalid_scope|Request had insufficient authentication scopes|403/i.test(msg);
}

const b64 = (s: string) => Buffer.from(s, 'utf-8').toString('base64');

/** 텍스트 메일 1통 (제목·본문 한글 안전 — UTF-8 base64 인코딩) */
export async function sendMail(params: { to: string; subject: string; text: string }): Promise<MailResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };

    const mime = [
      `To: ${params.to}`,
      `Subject: =?UTF-8?B?${b64(params.subject)}?=`,
      'MIME-Version: 1.0',
      'Content-Type: text/plain; charset=UTF-8',
      'Content-Transfer-Encoding: base64',
      '',
      b64(params.text),
    ].join('\r\n');
    const raw = Buffer.from(mime, 'utf-8').toString('base64url');

    const gmail = google.gmail({ version: 'v1', auth });
    await gmail.users.messages.send({ userId: 'me', requestBody: { raw } });
    return { ok: true };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, error: msg, needsReauth: isScopeError(msg), apiDisabled: isApiDisabled(msg) };
  }
}
