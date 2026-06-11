/**
 * 시리얼 번호 형식 유틸 — 클라이언트/서버 공용 (순수 함수, DB 의존 없음).
 *
 * 신규 형식: M{YY}-{NNNN}  (예: M26-0042 = 2026년 42번째)
 *  - M = MAMORU, YY = 생성연도 2자리, NNNN = 연도내 전역 일련(4자리 패딩)
 *  - 각인 시 하이픈 생략 가능(M260042) → 조회는 normalizeSerial로 하이픈/대소문자 무시
 *  - 기존 8자리 숫자 시리얼(13790001 등)은 이 형식과 무관하게 보존
 */

const SERIAL_RE = /^M(\d{2})-?(\d{4,})$/i;

/** M{YY}-{NNNN} 문자열 생성 */
export function formatSerial(year2: number, seq: number): string {
  return `M${String(year2).padStart(2, '0')}-${String(seq).padStart(4, '0')}`;
}

/** M{YY}-{NNNN}(하이픈 유무 무관) 파싱 → {year2, seq} 또는 null */
export function parseSerial(s: string): { year2: number; seq: number } | null {
  const m = String(s ?? '').trim().match(SERIAL_RE);
  if (!m) return null;
  return { year2: parseInt(m[1], 10), seq: parseInt(m[2], 10) };
}

/** 다음 시리얼 문자열 (M형식이면 seq+1, 아니면 원본 반환) */
export function incrementSerial(s: string): string {
  const p = parseSerial(s);
  if (!p) return s;
  return formatSerial(p.year2, p.seq + 1);
}

/** 조회/중복 비교용 정규화 — 하이픈 제거 + 대문자 */
export function normalizeSerial(s: string): string {
  return String(s ?? '').trim().toUpperCase().replace(/-/g, '');
}

/** 현재 연도 2자리 (예: 2026 → 26) */
export function currentYear2(date: Date = new Date()): number {
  return date.getFullYear() % 100;
}

/** M형식 시리얼인지 */
export function isMSerial(s: string): boolean {
  return parseSerial(s) !== null;
}
