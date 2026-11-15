/**
 * 검색 맞추기 한 곳 (2026-11-15) — 검색해서 고르는 칸은 모두 이 규칙으로 걸러진다
 *
 *   · 대소문자 · 띄어쓰기 · 하이픈 무시        "bl 600" = "BL-600"
 *   · 전화번호는 숫자만 비교(뒤 4자리로도)      "5678" → 010-1234-5678
 *   · 초성으로도                                 "ㅇㅍ" → 알파 블런트
 *
 * 쓰는 법:  matchSearch(query, product.name, product.sku)   // 뒤에 넣은 글자 중 하나라도 맞으면 true
 */

const CHO = ['ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ', 'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ'];

/** 비교용으로 다듬기: 소문자 · 띄어쓰기/하이픈/점/괄호 제거 */
export function normalizeSearch(v: unknown): string {
  return String(v ?? '').toLowerCase().replace(/[\s\-_.()/·]/g, '');
}

/** 한글 글자마다 초성만 뽑기 ("알파 블런트" → "ㅇㅍㅂㄹㅌ"). 한글이 아닌 글자는 그대로 둔다 */
export function toChosung(v: unknown): string {
  let out = '';
  for (const ch of normalizeSearch(v)) {
    const code = ch.charCodeAt(0);
    if (code >= 0xac00 && code <= 0xd7a3) out += CHO[Math.floor((code - 0xac00) / 588)];
    else out += ch;
  }
  return out;
}

const ONLY_CHO = /^[ㄱ-ㅎ]+$/;

/** query 가 texts 중 하나에라도 맞으면 true. query 가 비어 있으면 항상 true */
export function matchSearch(query: string, ...texts: Array<string | null | undefined>): boolean {
  const q = normalizeSearch(query);
  if (!q) return true;
  const byCho = ONLY_CHO.test(q);
  for (const t of texts) {
    if (!t) continue;
    const n = normalizeSearch(t);
    if (n.includes(q)) return true;
    if (byCho && toChosung(t).includes(q)) return true;
  }
  return false;
}
