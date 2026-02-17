/**
 * 상담 관련 타입 정의
 * GAS Push 방식이므로 GAS 스크립트 측에서 데이터를 가공해서 전송
 */

// GasPushPayload는 sync.ts에서 직접 정의/export
// 이 파일은 추가 타입이 필요할 때 확장용으로 유지

/** 상태 전이 허용 맵 */
export const VALID_STATUS_TRANSITIONS: Record<string, string[]> = {
  pending_admin: ['suggested', 'assigned', 'cancelled'],
  suggested: ['confirmed', 'reschedule_requested', 'cancelled'],
  assigned: ['confirmed', 'reschedule_requested', 'cancelled'],
  reschedule_requested: ['suggested', 'assigned', 'cancelled'],
  confirmed: [],
  cancelled: [],
};
