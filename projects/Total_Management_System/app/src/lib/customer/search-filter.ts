/**
 * 고객 검색 조건 (2026-11-20) — 고객 목록 · 고객 자동완성이 같은 식을 쓴다
 *
 * 찾는 칸: 고객명 · 업체/매장명 · 활동명 · 상세 주소 · (숫자가 2자리 이상이면) 연락처
 * - 상세 주소: 매장 이름을 주소 칸에 적어 둔 고객이 있어서 넣었다(매장명으로 쳐도 안 나오던 문제).
 * - 도로명 주소는 넣지 않는다 — "서울" · "강남"처럼 흔한 말에 고객이 줄줄이 걸려 정작 찾는 사람이 밀린다.
 * - 두 글자부터 찾는다. 한 글자는 걸리는 고객이 너무 많다.
 */
export const CUSTOMER_SEARCH_MIN = 2;

/** 검색어 다듬기 — 조건식을 깨는 글자(% , ( ))를 뺀다 */
export function cleanCustomerSearch(raw: string | null | undefined): string {
  return (raw || '').replace(/[%,()]/g, '').trim();
}

/** PostgREST or 조건. 두 글자가 안 되면 null(검색하지 않음) */
export function customerSearchOr(raw: string | null | undefined): string | null {
  const q = cleanCustomerSearch(raw);
  if (q.length < CUSTOMER_SEARCH_MIN) return null;
  const parts = [
    `name.ilike.%${q}%`,
    `company_name.ilike.%${q}%`,
    `activity_name.ilike.%${q}%`,
    `address_detail.ilike.%${q}%`,
  ];
  const digits = q.replace(/\D/g, '');
  if (digits.length >= 2) parts.push(`phone_normalized.ilike.%${digits}%`);
  return parts.join(',');
}
