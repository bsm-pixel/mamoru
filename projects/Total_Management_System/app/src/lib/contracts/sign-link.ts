/**
 * 계약서 고객 서명 링크 (2026-11-22, 마이그 172)
 *
 * 온라인(톡상담) 고객에게 계약서를 링크로 보내 폰에서 읽고 서명하게 한다.
 * 고객 화면 = projects/contract/page_sign.html (GitHub Pages) → /api/contracts/public 로 읽고 쓴다.
 * 링크 열쇠 = contracts.sign_token (추측 불가 uuid). 이름·전화·주소는 고객 화면에 보내지 않는다.
 */

export const SIGN_PAGE_BASE = 'https://page.mamoru.kr/projects/contract/page_sign.html';

export function signUrl(token: string): string {
  return `${SIGN_PAGE_BASE}?t=${encodeURIComponent(token)}`;
}

/** 톡에 붙여 넣을 문구 */
export function signMessage(name: string, token: string, opts?: { engravingOnly?: boolean }): string {
  const who = (name || '').trim() ? `${name.trim()}님, ` : '';
  if (opts?.engravingOnly) {
    return `[MAMORU] 이니셜 각인 문구\n${who}각인할 문구가 정해지시면 아래 주소에서 보내 주세요.\n${signUrl(token)}`;
  }
  return `[MAMORU] 구매 계약서\n${who}아래 주소에서 계약 내용을 확인하시고 서명해 주세요.\n${signUrl(token)}`;
}

export const ENGRAVING_MAX = 30;

/** 각인 문구 다듬기 — 앞뒤 공백·줄바꿈 정리, 최대 30자. 비면 '' */
export function cleanEngraving(v: unknown): string {
  return String(v ?? '').replace(/\s+/g, ' ').trim().slice(0, ENGRAVING_MAX);
}

/** 서명 그림 검사 — 캔버스에서 만든 PNG data URL 만, 400KB 이하 */
export function isValidSignature(v: unknown): v is string {
  return typeof v === 'string' && v.startsWith('data:image/png;base64,') && v.length > 200 && v.length < 400_000;
}

/** 고객 화면에 보낼 이름 — 가운데 글자를 가린다 (링크가 다른 사람에게 넘어가도 실명이 다 보이지 않게) */
export function maskName(n: string | null | undefined): string {
  const s = (n || '').trim();
  if (s.length <= 1) return s;
  if (s.length === 2) return s[0] + '*';
  return s[0] + '*'.repeat(s.length - 2) + s[s.length - 1];
}

/** uuid 모양인지 (주소로 들어온 열쇠를 데이터베이스에 묻기 전에 거른다) */
export function isToken(v: unknown): v is string {
  return typeof v === 'string' && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(v);
}
