/**
 * Gmail 발송 (2026-09-29) — 관리자 푸시 미수신 시 사장님께 알림 메일
 *
 * 기존 구글 연결(캘린더·할 일과 같은 토큰)을 그대로 쓴다. `gmail.send` 스코프가 필요 → 설정에서 재연결 1회.
 * 사장님 계정 → 사장님 계정("나에게 보내기")이어도 Gmail 앱 알림이 울리는 것을 사장님이 직접 확인함(2026-09-29).
 * 실패는 throw 하지 않고 결과로 반환한다 — 호출처가 기록한다.
 * 2026-10-29: 읽기(listMails) 추가 — 영수증 메일 → 경비 확인 대기. `gmail.readonly` 필요.
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

/* ── 읽기 (2026-10-29) — 영수증·세금계산서 메일 → 경비 확인 대기 (lib/expenses/mail-scan.ts) ──
   `gmail.readonly` 스코프 필요 → 설정에서 재연결 1회. 실패는 결과로 반환(호출처가 화면에 안내). */

export interface MailMessage {
  id: string;
  from: string;          // "Adobe <pgadmcust@kcp.co.kr>" 원문
  subject: string;
  receivedAt: string;    // ISO
  text: string;          // 본문(텍스트 우선, 없으면 HTML 을 글자만 남김)
}

export interface MailListResult extends Omit<MailResult, 'ok'> {
  ok: boolean;
  messages: MailMessage[];
}

type Part = { mimeType?: string | null; body?: { data?: string | null } | null; parts?: Part[] | null };

function decodeB64Url(data: string): string {
  return Buffer.from(data.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf-8');
}

/** HTML → 글자만 (표 칸은 " | " 로 남겨 "결제금액 | 30,500원" 같은 짝이 붙어 있게) */
export function htmlToText(html: string): string {
  return html
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<\/(td|th)>/gi, ' | ')
    .replace(/<br\s*\/?>|<\/(p|div|tr|li|h\d)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#8361;|&#x20a9;/gi, '₩')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n');
}

function extractText(part: Part | undefined | null): { plain: string; html: string } {
  const out = { plain: '', html: '' };
  const walk = (p: Part | undefined | null) => {
    if (!p) return;
    const data = p.body?.data;
    if (data && p.mimeType === 'text/plain') out.plain += decodeB64Url(data) + '\n';
    else if (data && p.mimeType === 'text/html') out.html += decodeB64Url(data) + '\n';
    (p.parts || []).forEach(walk);
  };
  walk(part);
  return out;
}

/** 검색어로 메일 목록 + 본문 (최대 max 통, exclude 에 있는 id 는 건너뜀) */
export async function listMails(query: string, max = 100, exclude?: Set<string>): Promise<MailListResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected', messages: [] };
    const gmail = google.gmail({ version: 'v1', auth });

    const ids: string[] = [];
    let pageToken: string | undefined;
    do {
      const res = await gmail.users.messages.list({ userId: 'me', q: query, maxResults: Math.min(100, max - ids.length), pageToken });
      for (const m of res.data.messages || []) if (m.id) ids.push(m.id);
      pageToken = res.data.nextPageToken || undefined;
    } while (pageToken && ids.length < max);

    const messages: MailMessage[] = [];
    for (const id of ids) {
      if (exclude?.has(id)) continue; // 이미 처리한 메일은 본문을 다시 받지 않는다
      const { data } = await gmail.users.messages.get({ userId: 'me', id, format: 'full' });
      const headers = data.payload?.headers || [];
      const h = (name: string) => headers.find((x) => (x.name || '').toLowerCase() === name)?.value || '';
      const { plain, html } = extractText(data.payload as Part);
      messages.push({
        id,
        from: h('from'),
        subject: h('subject'),
        receivedAt: new Date(Number(data.internalDate || Date.now())).toISOString(),
        // 카드사 메일은 text/plain 이 비었거나 깨진 경우가 많다 → HTML 이 있으면 HTML 을 기준으로
        text: html ? htmlToText(html) : plain,
      });
    }
    return { ok: true, messages };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, error: msg, needsReauth: isScopeError(msg), apiDisabled: isApiDisabled(msg), messages: [] };
  }
}
