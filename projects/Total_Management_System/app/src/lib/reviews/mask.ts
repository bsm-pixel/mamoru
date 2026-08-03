/**
 * 리뷰 이벤트 공개 노출용 마스킹 유틸 (서버 전용).
 * 공개 API 는 실명/전화 원본을 절대 내보내지 않는다 — 반드시 이 함수를 거친 값만 반환.
 * 표기 규칙은 고객 페이지(page_review_event.html) 시안과 일치: 홍**님 / 010-****-32**
 */

/** 이름 마스킹: 첫 글자만 노출, 나머지는 * (홍길동→홍**, 김미→김*). '님'은 호출부에서 붙임 */
export function maskNameEvent(name: string | null | undefined): string {
  const n = (name ?? '').trim();
  if (!n) return '고객';
  if (n.length === 1) return n;
  return n[0] + '*'.repeat(n.length - 1);
}

/** 전화 마스킹: 앞 3자리 + 가운데 전부 * + 끝 4자리 중 앞 2자리만 (01012343288 → 010-****-32**) */
export function maskPhoneEvent(phone: string | null | undefined): string {
  const d = (phone ?? '').replace(/\D/g, '');
  if (d.length < 9) return ''; // 형식 미달이면 노출 안 함
  const head = d.slice(0, 3);
  const last4 = d.slice(-4);
  return `${head}-****-${last4.slice(0, 2)}**`;
}

/** 표시명(오버라이드 우선) + 님. 예: '홍**님' */
export function displayWinnerName(rawName: string | null | undefined, override?: string | null): string {
  const ov = (override ?? '').trim();
  if (ov) return ov;
  return maskNameEvent(rawName) + '님';
}
