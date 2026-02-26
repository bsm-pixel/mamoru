/**
 * 복원수리 상태 전이 규칙
 * 서버(API)와 클라이언트(UI) 모두에서 공유
 */

import type { RepairStatus } from '@/lib/supabase/types';

type TransitionMap = Record<string, RepairStatus[]>;

/** 복원수리 상태 전이 맵 (v3 — payment_confirmed 파이프라인 분리) */
const REPAIR_TRANSITIONS: TransitionMap = {
  intake:             ['pickup_scheduled', 'cost_notified', 'cancelled'],
  pickup_scheduled:   ['cost_notified', 'cancelled'],
  cost_notified:      ['repairing', 'cancelled'],   // payment_confirmed 제거
  repairing:          ['ready_to_ship', 'cancelled'],  // R1: ready_to_ship 활성 승격
  shipped:            ['delivered'],
  delivered:          ['completed'],
  completed:          [],
  cancelled:          [],
  ready_to_ship:      ['shipped'],  // R1: 활성 상태로 승격
  // 레거시 호환 (기존 데이터)
  picked_up:          ['cost_notified', 'cancelled'],
  inspecting:         ['cost_notified', 'cancelled'],
  payment_confirmed:  ['repairing', 'cancelled'],    // 레거시 데이터 전이용
};

/** 주어진 상태에서 전이 가능한 상태 목록 */
export function getAllowedRepairTransitions(currentStatus: RepairStatus): RepairStatus[] {
  return REPAIR_TRANSITIONS[currentStatus] || [];
}

/** 전이 가능 여부 검증 */
export function isValidRepairTransition(from: RepairStatus, to: RepairStatus): boolean {
  return getAllowedRepairTransitions(from).includes(to);
}

/** 상태 순서 (진행도 계산용) — payment_confirmed 제거 */
export const REPAIR_STATUS_ORDER: RepairStatus[] = [
  'intake',
  'pickup_scheduled',
  'cost_notified',
  'repairing',
  'ready_to_ship',  // R1: 출고대기 활성 추가
  'shipped',
  'delivered',
  'completed',
];

/** 상태 한글 라벨 */
export const REPAIR_STATUS_LABEL: Record<string, string> = {
  intake: '신규접수',
  pickup_scheduled: '입고대기',
  cost_notified: '작업중',
  payment_confirmed: '작업중',  // 레거시
  repairing: '작업중',
  ready_to_ship: '출고대기',    // R1: 활성 상태로 승격
  shipped: '출고완료',
  delivered: '배송완료',
  completed: '완료',
  cancelled: '취소',
  // 레거시
  picked_up: '입고완료',
  inspecting: '검수중',
};

/** 상태 색상 (Tailwind 클래스) */
export const REPAIR_STATUS_COLOR: Record<string, string> = {
  intake: 'bg-info-soft text-info',
  pickup_scheduled: 'bg-info-soft text-info',
  cost_notified: 'bg-terracotta-soft/30 text-terracotta-deep',
  payment_confirmed: 'bg-terracotta-soft/30 text-terracotta-deep',
  repairing: 'bg-terracotta-soft/30 text-terracotta-deep',
  shipped: 'bg-success-soft text-success',
  delivered: 'bg-success-soft text-success',
  completed: 'bg-neutral-100 text-neutral-500',
  cancelled: 'bg-error-soft text-error',
  // 레거시
  picked_up: 'bg-info-soft text-info',
  inspecting: 'bg-warning-soft text-warning',
  ready_to_ship: 'bg-success-soft text-success',
};

/** 상태 전이 → 버튼 라벨 매핑 */
export const REPAIR_ACTION_LABEL: Record<string, string> = {
  intake: '접수',
  pickup_scheduled: '수거접수 완료',
  cost_notified: '입고 & 비용안내',
  repairing: '작업 시작',
  shipped: '출고',
  delivered: '배송완료',
  completed: '완료',
  cancelled: '취소',
  // 레거시
  picked_up: '입고완료',
  inspecting: '검수 시작',
  ready_to_ship: '출고대기',
};

/** 진행방식별 분기 라벨 (방문수거 intake → "수거접수 필요") */
export function getRepairDisplayLabel(status: RepairStatus, proceedType?: string | null): string {
  if (status === 'intake' && proceedType === '방문수거') return '수거접수 필요';
  return REPAIR_STATUS_LABEL[status] || status;
}

/** 진행방식별 허용 전이 필터 */
export function getFilteredRepairTransitions(
  status: RepairStatus,
  proceedType?: string | null
): RepairStatus[] {
  const all = getAllowedRepairTransitions(status);
  if (status === 'intake' && proceedType === '방문수거') {
    return all.filter(s => s !== 'cost_notified'); // 방문수거: 수거접수 먼저
  }
  if (status === 'intake' && proceedType !== '방문수거') {
    return all.filter(s => s !== 'pickup_scheduled'); // 직접발송: 수거접수 스킵
  }
  return all;
}
