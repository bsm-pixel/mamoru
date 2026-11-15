/**
 * 고객 숨기기 (2026-11-15) — 삭제 규칙 통일: "거래 · 이력이 없으면 삭제, 있으면 숨기기(되살릴 수 있음)"
 *
 * 거래가 있는 고객은 지울 수 없다(판매 · 복원수리 기록이 고객을 가리킨다). 대신 목록 · 검색에서 뺀다.
 * 표에 칸을 새로 만들지 않고 고객 꼬리표(tags)에 「숨김」을 붙이는 방식 — 되살리려면 꼬리표만 떼면 된다.
 */
export const HIDDEN_TAG = '숨김';

export function isHiddenCustomer(tags: string[] | null | undefined): boolean {
  return Array.isArray(tags) && tags.includes(HIDDEN_TAG);
}

/** PostgREST or 필터: 숨김 꼬리표가 없는 고객만 (꼬리표가 비어 있는 고객 포함) */
export const NOT_HIDDEN_FILTER = `tags.is.null,tags.not.cs.{${HIDDEN_TAG}}`;
