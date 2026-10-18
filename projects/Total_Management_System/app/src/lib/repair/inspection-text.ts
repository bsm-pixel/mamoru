/**
 * 검수 결과 → 자동 문구 생성
 * 2026-05-29: 7항목 체크리스트 → 사진 위 핀 마킹(photo_marks) 기반으로 전환.
 *   - 핀 라벨(무뎌짐/찍힘/빗살 손상/날 보정/장력 조정/부품 교체/스토퍼 교체 등)을 집계해 요약 생성.
 *   - 고객 안내문 본문은 repairs.admin_note (멘트 프리셋) 가 담당.
 * 2026-10-15 v3: diagram_marks(그림체 체크)가 있으면 그것 기준 — 문구는 wound-catalog SSOT
 */

import type { RepairInspection } from '@/lib/supabase/types';
import { FLAG_TYPES } from '@/components/repairs/inspection-marks';
import { WOUND_KINDS, checkIdsOf, normalizeChecks, woundSummary } from '@/lib/repair/wound-catalog';

type Mark = { x?: number; y?: number; label: string; flag?: boolean };

// 2026-10-15: 플래그(장력·밸런스·날각 — 위치 없음, label=긴 안내문)는 "N곳" 집계 대상이 아니다
//   (전엔 "장력이 헐거운 상태 - 조절 필요 1곳"처럼 TMS 요약·고객 work_summary에 섞여 나왔다)
function marksOf(insp: RepairInspection): Mark[] {
  return ((insp.photo_marks as Mark[] | null) || []).filter((m) => !m.flag);
}

/** 가위 한 자루 요약 — 핀 라벨별 개수 (예: "무뎌짐 2 · 찍힘") */
export function getScissorSummary(insp: RepairInspection): string {
  if (insp.diagram_marks != null) return woundSummary(normalizeChecks(insp.diagram_marks));
  const marks = marksOf(insp);
  // 플래그는 개수 없이 짧은 라벨로 덧붙인다 (저장값=긴 note → FLAG_TYPES 로 짧은 라벨 역매핑)
  const flags = ((insp.photo_marks as Mark[] | null) || [])
    .filter((m) => m.flag)
    .map((m) => FLAG_TYPES.find((f) => f.note === m.label)?.label || m.label);
  if (marks.length === 0 && flags.length === 0) return '표시 없음';
  const counts: Record<string, number> = {};
  for (const m of marks) counts[m.label] = (counts[m.label] || 0) + 1;
  return [
    ...Object.entries(counts).map(([label, n]) => (n > 1 ? `${label} ${n}` : label)),
    ...flags,
  ].join(' · ');
}

/** 종합 자동 문구 — 전체 핀 라벨 집계 (예: "무뎌짐 3곳 · 찍힘 1곳 복원") */
export function generateWorkSummary(inspections: RepairInspection[]): string {
  if (inspections.length === 0) return '';
  const counts: Record<string, number> = {};
  const v3: Record<string, number> = {}; // v3: 항목별 해당 가위 수
  for (const insp of inspections) {
    if (insp.diagram_marks != null) {
      const checks = normalizeChecks(insp.diagram_marks);
      for (const k of WOUND_KINDS) if (checkIdsOf(k).some((id) => checks.includes(id))) v3[k.ko] = (v3[k.ko] || 0) + 1;
      continue;
    }
    for (const m of marksOf(insp)) counts[m.label] = (counts[m.label] || 0) + 1;
  }
  const parts = [
    ...Object.entries(counts).map(([label, n]) => `${label} ${n}곳`),
    ...Object.entries(v3).map(([label, n]) => `${label} ${n}자루`),
  ];
  if (parts.length === 0) return '';
  return parts.join(' · ') + ' 복원';
}
