/**
 * 복원수리 상태 전이 규칙
 * 서버(API)와 클라이언트(UI) 모두에서 공유
 */

import type { RepairStatus } from '@/lib/supabase/types';

type TransitionMap = Record<string, RepairStatus[]>;

/** 복원수리 상태 전이 맵 */
const REPAIR_TRANSITIONS: TransitionMap = {
  intake:             ['pickup_scheduled', 'picked_up', 'inspecting', 'cancelled'],
  pickup_scheduled:   ['picked_up', 'cancelled'],
  picked_up:          ['inspecting', 'cancelled'],
  inspecting:         ['cost_notified', 'cancelled'],
  cost_notified:      ['payment_confirmed', 'cancelled'],
  payment_confirmed:  ['repairing', 'cancelled'],
  repairing:          ['ready_to_ship'],
  ready_to_ship:      ['shipped'],
  shipped:            ['delivered'],
  delivered:          ['completed'],
  completed:          [],
  cancelled:          [],
};

/** 주어진 상태에서 전이 가능한 상태 목록 */
export function getAllowedRepairTransitions(currentStatus: RepairStatus): RepairStatus[] {
  return REPAIR_TRANSITIONS[currentStatus] || [];
}

/** 전이 가능 여부 검증 */
export function isValidRepairTransition(from: RepairStatus, to: RepairStatus): boolean {
  return getAllowedRepairTransitions(from).includes(to);
}

/** 상태 순서 (진행도 계산용) */
export const REPAIR_STATUS_ORDER: RepairStatus[] = [
  'intake',
  'pickup_scheduled',
  'picked_up',
  'inspecting',
  'cost_notified',
  'payment_confirmed',
  'repairing',
  'ready_to_ship',
  'shipped',
  'delivered',
  'completed',
];

/** 상태 한글 라벨 */
export const REPAIR_STATUS_LABEL: Record<RepairStatus, string> = {
  intake: '접수',
  pickup_scheduled: '수거예약',
  picked_up: '수거완료',
  inspecting: '검수중',
  cost_notified: '비용안내',
  payment_confirmed: '입금확인',
  repairing: '수리중',
  ready_to_ship: '출고대기',
  shipped: '발송완료',
  delivered: '배송완료',
  completed: '완료',
  cancelled: '취소',
};

/** 상태 색상 (Tailwind 클래스) */
export const REPAIR_STATUS_COLOR: Record<RepairStatus, string> = {
  intake: 'bg-info-soft text-info',
  pickup_scheduled: 'bg-info-soft text-info',
  picked_up: 'bg-info-soft text-info',
  inspecting: 'bg-warning-soft text-warning',
  cost_notified: 'bg-warning-soft text-warning',
  payment_confirmed: 'bg-terracotta-soft/30 text-terracotta-deep',
  repairing: 'bg-terracotta-soft/30 text-terracotta-deep',
  ready_to_ship: 'bg-success-soft text-success',
  shipped: 'bg-success-soft text-success',
  delivered: 'bg-success-soft text-success',
  completed: 'bg-neutral-100 text-neutral-500',
  cancelled: 'bg-error-soft text-error',
};

/** 상태 전이 → 버튼 라벨 매핑 */
export const REPAIR_ACTION_LABEL: Record<RepairStatus, string> = {
  intake: '접수',
  pickup_scheduled: '수거예약',
  picked_up: '수거완료',
  inspecting: '검수 시작',
  cost_notified: '비용 안내',
  payment_confirmed: '입금 확인',
  repairing: '수리 시작',
  ready_to_ship: '수리 완료',
  shipped: '출고',
  delivered: '배송완료',
  completed: '완료',
  cancelled: '취소',
};
