/**
 * 리뷰 이벤트 공개 노출용 마스킹 유틸 (서버·클라 공용 — 순수 함수).
 * 공개 API 는 실명/전화 원본을 절대 내보내지 않는다 — 반드시 이 함수를 거친 값만 반환.
 * 표기 규칙(고객 페이지 시안 일치): 이름=성 가운데 * (백성민→백*민) / 전화=뒷 4자리 (…3562→(3562))
 * 예) 백성민 01022483562 → "백*민 님 (3562)"
 */

/** 이름 마스킹: 첫·끝 글자만 노출, 가운데는 * (백성민→백*민, 남궁민수→남**수, 김민→김*, 김→김). '님'은 displayWinnerName 에서 붙임 */
export function maskNameEvent(name: string | null | undefined): string {
  const n = (name ?? '').trim();
  if (!n) return '고객';
  if (n.length === 1) return n;
  if (n.length === 2) return n[0] + '*';
  return n[0] + '*'.repeat(n.length - 2) + n[n.length - 1];
}

/** 전화 마스킹: 뒷 4자리만 괄호로 (01022483562 → "(3562)"). 자릿수 미달이면 노출 안 함 */
export function maskPhoneEvent(phone: string | null | undefined): string {
  const d = (phone ?? '').replace(/\D/g, '');
  if (d.length < 4) return '';
  return `(${d.slice(-4)})`;
}

/** 표시명(오버라이드 우선) + 님. 예: '백*민 님' (오버라이드 있으면 그대로) */
export function displayWinnerName(rawName: string | null | undefined, override?: string | null): string {
  const ov = (override ?? '').trim();
  if (ov) return ov;
  return maskNameEvent(rawName) + ' 님';
}
