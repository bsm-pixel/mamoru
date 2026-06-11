/**
 * 시리얼 번호 형식 유틸 — 클라이언트/서버 공용 (순수 함수, DB 의존 없음).
 *
 * 형식: MR{YY}{NNNNN}  (예: MR2610816)
 *  - MR = MAMORU, YY = 생성연도 2자리(제조/판매 시점 표기), NNNNN = 전역 누적 일련번호(5자리, 리셋 없음)
 *  - 누적 번호는 연도가 바뀌어도 0으로 리셋하지 않고 계속 증가(2027년도 이어서) → "갓 시작" 느낌·판매량 노출 방지
 *  - 하이픈 없이 9자리 고정. 99999 초과 시 자연히 6자리로 확장
 *  - 기존 8자리 숫자 시리얼(13790001 등)·구 M26- 형식은 이 형식과 무관하게 보존(파싱 null)
 */

const SERIAL_RE = /^MR(\d{2})(\d{5,})$/i;

/** MR{YY}{NNNNN} 문자열 생성 (n = 전역 누적 번호) */
export function formatSerial(year2: number, n: number): string {
  return `MR${String(year2).padStart(2, '0')}${String(n).padStart(5, '0')}`;
}

/** MR{YY}{NNNNN} 파싱 → {year2, seq(=누적번호)} 또는 null */
export function parseSerial(s: string): { year2: number; seq: number } | null {
  const m = String(s ?? '').trim().match(SERIAL_RE);
  if (!m) return null;
  return { year2: parseInt(m[1], 10), seq: parseInt(m[2], 10) };
}

/** 다음 시리얼 문자열 (MR형식이면 누적번호+1, 아니면 원본 반환) */
export function incrementSerial(s: string): string {
  const p = parseSerial(s);
  if (!p) return s;
  return formatSerial(p.year2, p.seq + 1);
}

/** 조회/중복 비교용 정규화 — 하이픈·공백 제거 + 대문자 */
export function normalizeSerial(s: string): string {
  return String(s ?? '').trim().toUpperCase().replace(/[-\s]/g, '');
}

/** 현재 연도 2자리 (예: 2026 → 26) */
export function currentYear2(date: Date = new Date()): number {
  return date.getFullYear() % 100;
}

/** MR형식 시리얼인지 */
export function isMSerial(s: string): boolean {
  return parseSerial(s) !== null;
}
