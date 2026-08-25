/**
 * 일정 카테고리 색상 — 단일 진실점(SSOT), 2026-08-12
 *   인앱 달력(대시보드·일정·상담 매장탭)과 구글 캘린더 colorId를 한곳에서 정의.
 *   기존에 매장/출장 색이 인앱↔구글, 대시보드↔매장탭 사이에서 어긋났던 문제 해결.
 *
 * 매장=초록(emerald) / 출장=보라(violet) / 수리 방문=주황(amber) / 수거=파랑(sky)
 *
 * ⚠️ Tailwind는 이 파일의 문자열 리터럴을 스캔해 클래스를 생성하므로,
 *    dot/badge/text 값은 반드시 완성된 클래스명 리터럴로 유지할 것.
 * Google colorId 표: https://developers.google.com/calendar/api/v3/reference/colors
 */

export type ScheduleCategory = 'store' | 'field' | 'repair_visit' | 'repair_pickup' | 'return_pickup';

export interface ScheduleColor {
  label: string;
  dot: string;          // 달력 점(기본)
  dotSelected: string;  // 선택된(어두운) 날짜 위의 밝은 점
  badge: string;        // 목록 배지 (bg + text)
  text: string;         // 아이콘/텍스트 색
  googleColorId: string | null; // 구글 캘린더 colorId (null = 구글 미동기화)
}

export const SCHEDULE_COLORS: Record<ScheduleCategory, ScheduleColor> = {
  store: {
    label: '매장',
    dot: 'bg-emerald-500', dotSelected: 'bg-emerald-300',
    badge: 'bg-emerald-50 text-emerald-700', text: 'text-emerald-600',
    googleColorId: '2', // Sage (녹색)
  },
  field: {
    label: '출장',
    dot: 'bg-violet-500', dotSelected: 'bg-violet-300',
    badge: 'bg-violet-50 text-violet-700', text: 'text-violet-600',
    googleColorId: '3', // Grape (보라)
  },
  repair_visit: {
    label: '수리',
    dot: 'bg-amber-500', dotSelected: 'bg-amber-300',
    badge: 'bg-amber-50 text-amber-700', text: 'text-amber-600',
    googleColorId: '6', // Tangerine (주황)
  },
  repair_pickup: {
    label: '수거',
    dot: 'bg-sky-500', dotSelected: 'bg-sky-300',
    badge: 'bg-sky-50 text-sky-700', text: 'text-sky-600',
    googleColorId: null, // 방문수거는 구글 캘린더 미동기화(직접방문만 동기화)
  },
  return_pickup: {
    label: '반품수거',
    dot: 'bg-rose-500', dotSelected: 'bg-rose-300',
    badge: 'bg-rose-50 text-rose-700', text: 'text-rose-600',
    googleColorId: null, // 반품수거는 인앱 달력만
  },
};
