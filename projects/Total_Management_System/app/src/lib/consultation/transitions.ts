/**
 * 상담 유형별 상태 전이 규칙
 * 서버(API)와 클라이언트(UI) 모두에서 공유
 */

import type { ConsultationStatus, ConsultationType } from '@/lib/supabase/types';

type TransitionMap = Record<string, ConsultationStatus[]>;

/** 매장방문 상태 전이 */
const STORE_VISIT: TransitionMap = {
  confirmed: ['completed', 'on_hold', 'cancelled'],
  on_hold: ['confirmed', 'cancelled'],
};

/** 출장요청 상태 전이 */
const FIELD_REQUEST: TransitionMap = {
  pending_admin: ['suggested', 'confirmed', 'on_hold', 'cancelled'],
  suggested: ['confirmed', 'reschedule_requested', 'on_hold', 'cancelled'],
  reschedule_requested: ['suggested', 'confirmed', 'on_hold', 'cancelled'],
  change_requested: ['suggested', 'confirmed', 'on_hold', 'cancelled'],  // 고객 일정변경 요청
  confirmed: ['completed', 'on_hold', 'cancelled'],
  on_hold: ['pending_admin', 'suggested', 'confirmed', 'cancelled'],
};

/** 톡상담 상태 전이 */
const TALK_CONSULT: TransitionMap = {
  pending_admin: ['in_progress', 'on_hold', 'cancelled'],
  in_progress: ['completed', 'on_hold', 'cancelled'],
  on_hold: ['pending_admin', 'in_progress', 'cancelled'],
  // completed, cancelled → 종료 상태
};

const TYPE_MAP: Record<ConsultationType, TransitionMap> = {
  store_visit: STORE_VISIT,
  field_request: FIELD_REQUEST,
  talk_consult: TALK_CONSULT,
};

/** 주어진 유형·현재 상태에서 전이 가능한 상태 목록 반환 */
export function getAllowedTransitions(
  type: ConsultationType,
  currentStatus: ConsultationStatus
): ConsultationStatus[] {
  return TYPE_MAP[type]?.[currentStatus] || [];
}

/** 전이 가능 여부 검증 */
export function isValidTransition(
  type: ConsultationType,
  from: ConsultationStatus,
  to: ConsultationStatus
): boolean {
  return getAllowedTransitions(type, from).includes(to);
}
