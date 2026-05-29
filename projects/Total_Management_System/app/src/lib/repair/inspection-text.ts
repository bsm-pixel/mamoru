/**
 * 검수 결과 → 자동 문구 생성
 * 2026-05-29: 7항목 체크리스트 → 사진 위 핀 마킹(photo_marks) 기반으로 전환.
 *   - 핀 라벨(무뎌짐/찍힘/빗살 손상/날 보정/장력 조정/부품 교체/스토퍼 교체 등)을 집계해 요약 생성.
 *   - 고객 안내문 본문은 repairs.admin_note (멘트 프리셋) 가 담당.
 */

import type { RepairInspection } from '@/lib/supabase/types';

type Mark = { x: number; y: number; label: string };

function marksOf(insp: RepairInspection): Mark[] {
  return (insp.photo_marks as Mark[] | null) || [];
}

/** 가위 한 자루 요약 — 핀 라벨별 개수 (예: "무뎌짐 2 · 찍힘") */
export function getScissorSummary(insp: RepairInspection): string {
  const marks = marksOf(insp);
  if (marks.length === 0) return '표시 없음';
  const counts: Record<string, number> = {};
  for (const m of marks) counts[m.label] = (counts[m.label] || 0) + 1;
  return Object.entries(counts)
    .map(([label, n]) => (n > 1 ? `${label} ${n}` : label))
    .join(' · ');
}

/** 종합 자동 문구 — 전체 핀 라벨 집계 (예: "무뎌짐 3곳 · 찍힘 1곳 복원") */
export function generateWorkSummary(inspections: RepairInspection[]): string {
  if (inspections.length === 0) return '';
  const counts: Record<string, number> = {};
  for (const insp of inspections) {
    for (const m of marksOf(insp)) counts[m.label] = (counts[m.label] || 0) + 1;
  }
  const parts = Object.entries(counts).map(([label, n]) => `${label} ${n}곳`);
  if (parts.length === 0) return '';
  return parts.join(' · ') + ' 복원';
}
