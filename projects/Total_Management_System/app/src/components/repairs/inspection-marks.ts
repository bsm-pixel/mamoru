/**
 * 수리내역서 핀 마킹 v2 — 공용 타입/유형/색상 (보드·오버레이·고객 미리보기 공유)
 */

/** 사진 위 마크. x,y(0~1) 필수. x2,y2 있으면 선(드래그 범위). */
export interface MarkV2 {
  label: string;
  x: number;
  y: number;
  x2?: number;
  y2?: number;
}

/** 위치형 문제 유형 (사진 위 핀) */
export const MARK_TYPES: { label: string; color: string; hint?: string }[] = [
  { label: '무뎌짐', color: '#D97706', hint: '드래그=범위' },
  { label: '찍힘', color: '#DC2626' },
  { label: '빗살 손상', color: '#7C3AED' },
  { label: '부품교체 필요', color: '#EA580C' },
  { label: '스토퍼 문제', color: '#475569' },
];

/** 우측 상단 플래그 (위치 없음 — 부위 특정 불가, 선택 시 진단멘트 자동삽입) */
export const FLAG_TYPES: { key: string; label: string; note: string }[] = [
  { key: 'tension', label: '장력조절 필요', note: '장력이 헐거운 상태 - 조절 필요' },
  { key: 'balance', label: '밸런스 불균형', note: '가위 밸런스 불균형 — 교정 필요' },
  { key: 'edgeangle', label: '날각 문제', note: '날각 문제 — 날등 각도 개선 필요' },
];

export const FLAG_COLOR = '#B45309';

export function colorOf(label: string): string {
  return MARK_TYPES.find((t) => t.label === label)?.color || '#1A1A1A';
}
