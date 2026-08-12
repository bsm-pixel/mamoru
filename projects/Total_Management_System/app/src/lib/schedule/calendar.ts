/**
 * 일정 달력 공용 유틸 — 여러 달력(대시보드·일정·상담 매장탭)이 공유 (DRY, 2026-08-12)
 * 기존 3곳에 복붙돼 있던 getCalendarDays/formatYYYYMMDD를 단일 소스로 통합.
 * ymd는 로컬(KST) 기준이라 toISOString UTC 밀림 없음.
 */

/** 해당 월(0-based)의 달력 그리드(앞 공백 null + 1~말일 + 뒤 7배수 채움) */
export function getCalendarDays(year: number, month: number): (number | null)[] {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  const startOffset = firstDay.getDay(); // 0=일
  const totalDays = lastDay.getDate();
  const days: (number | null)[] = [];
  for (let i = 0; i < startOffset; i++) days.push(null);
  for (let d = 1; d <= totalDays; d++) days.push(d);
  while (days.length % 7 !== 0) days.push(null);
  return days;
}

/** year, month(0-based), day → 'YYYY-MM-DD' (로컬/KST 안전) */
export function ymd(year: number, month: number, day: number): string {
  return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}
