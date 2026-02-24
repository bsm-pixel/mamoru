/**
 * 검수 결과 → 자동 문구 생성 (GAS getRepairReportData_ L1294~L1326 이식)
 */

import type { RepairInspection } from '@/lib/supabase/types';

/** 문제점 판정 결과 */
interface InspectionIssues {
  hasMuddyim: boolean;       // 무뎌짐
  hasJjikim: boolean;         // 찍힘
  hasCombDamage: boolean;     // 빗살 손상
  hasLooseTension: boolean;   // 장력 헐거움
  hasPartsReplace: boolean;   // 내부부품 교체
  hasStopperReplace: boolean; // 스토퍼 교체
}

/** 점검 데이터에서 문제점 추출 */
export function analyzeIssues(inspections: RepairInspection[]): InspectionIssues {
  const issues: InspectionIssues = {
    hasMuddyim: false,
    hasJjikim: false,
    hasCombDamage: false,
    hasLooseTension: false,
    hasPartsReplace: false,
    hasStopperReplace: false,
  };

  for (const insp of inspections) {
    const fields = [insp.blade_tip, insp.blade_mid, insp.blade_inner];
    for (const v of fields) {
      if (v === '무뎌짐') issues.hasMuddyim = true;
      if (v === '찍힘') issues.hasJjikim = true;
    }
    if (insp.comb && insp.comb !== '양호' && insp.comb !== '') {
      issues.hasCombDamage = true;
    }
    if (insp.tension === '헐거움' || insp.tension === '교체') {
      issues.hasLooseTension = true;
    }
    if (insp.parts === '교체') {
      issues.hasPartsReplace = true;
    }
    if (insp.stopper === '교체') {
      issues.hasStopperReplace = true;
    }
  }

  return issues;
}

/** 문제점 → 종합 수리내역 자동 문구 */
export function generateWorkSummary(inspections: RepairInspection[]): string {
  if (inspections.length === 0) return '';

  const issues = analyzeIssues(inspections);
  const lines: string[] = [];

  // 날 상태 문구
  if (issues.hasMuddyim && issues.hasJjikim) {
    lines.push('무뎌짐 및 찍힘으로 인한 날 상처 복원작업 진행');
  } else if (issues.hasMuddyim) {
    lines.push('무뎌짐 확인되어 날 복원 작업');
  } else if (issues.hasJjikim) {
    lines.push('찍힘으로 인한 날 손상 복원작업');
  }

  // 빗살 문구
  if (issues.hasCombDamage) {
    lines.push('빗살의 손상은 복구 어렵습니다. 블레이드 날 보정으로 보완');
  }

  // 장력 문구
  if (issues.hasLooseTension) {
    lines.push('장력조절상태가 헐거울 경우 정확한 컷감이 어렵습니다. 볼트부를 시계방향으로 조절하여 사용해주세요.');
  }

  // 내부부품 문구
  if (issues.hasPartsReplace) {
    lines.push('내부 부품 노후로 인해 새 부품으로 교체 & 장착');
  }

  // 스토퍼 문구
  if (issues.hasStopperReplace) {
    lines.push('스토퍼 부위 문제로 인해 새 스토퍼로 교체 & 장착');
  }

  return lines.join('\n');
}

/** 가위 한 자루의 검수 요약 (한 줄) */
export function getScissorSummary(insp: RepairInspection): string {
  const problems: string[] = [];
  if (insp.blade_tip !== '양호') problems.push(`날끝:${insp.blade_tip}`);
  if (insp.blade_mid !== '양호') problems.push(`날중간:${insp.blade_mid}`);
  if (insp.blade_inner !== '양호') problems.push(`날안쪽:${insp.blade_inner}`);
  if (insp.comb && insp.comb !== '양호' && insp.comb !== '') problems.push(`빗살:${insp.comb}`);
  if (insp.tension !== '양호') problems.push(`장력:${insp.tension}`);
  if (insp.parts !== '양호') problems.push(`내부부품:${insp.parts}`);
  if (insp.stopper !== '양호') problems.push(`스토퍼:${insp.stopper}`);

  if (problems.length === 0) return '이상 없음';
  return problems.join(', ');
}
