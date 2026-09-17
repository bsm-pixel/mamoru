// 반품·교환수거 상태 전이 SSOT (2026-08-25 Phase 2) — 복원수리 transitions 패턴 복제
import type { ReturnStatus } from '@/lib/supabase/types';

/** 허용 전이 맵 */
export const RETURN_TRANSITIONS: Record<ReturnStatus, ReturnStatus[]> = {
  requested: ['pickup_scheduled', 'inbound', 'cancelled'],   // 직접반납 등은 수거예약 건너뛰고 바로 입고 가능
  pickup_scheduled: ['inbound', 'cancelled'],
  inbound: ['inspected', 'completed', 'cancelled'],
  inspected: ['completed', 'cancelled'],
  completed: [],
  cancelled: [],
};

export function getAllowedReturnTransitions(status: ReturnStatus): ReturnStatus[] {
  return RETURN_TRANSITIONS[status] || [];
}
export function isValidReturnTransition(from: ReturnStatus, to: ReturnStatus): boolean {
  return (RETURN_TRANSITIONS[from] || []).includes(to);
}

/** 진행도 순서 (프로그레스 바용) */
export const RETURN_STATUS_ORDER: ReturnStatus[] = [
  'requested', 'pickup_scheduled', 'inbound', 'inspected', 'completed',
];

/** 한글 라벨 */
export const RETURN_STATUS_LABEL: Record<ReturnStatus, string> = {
  requested: '수거접수',
  pickup_scheduled: '수거예약',
  inbound: '입고완료',
  inspected: '검수완료',
  completed: '완료',
  cancelled: '취소',
};

/** 상태 색 (Tailwind) */
export const RETURN_STATUS_COLOR: Record<ReturnStatus, string> = {
  requested: 'bg-amber-100 text-amber-700',
  pickup_scheduled: 'bg-blue-100 text-blue-700',
  inbound: 'bg-purple-100 text-purple-700',
  inspected: 'bg-indigo-100 text-indigo-700',
  completed: 'bg-emerald-100 text-emerald-700',
  cancelled: 'bg-red-100 text-red-700',
};

/** 액션 버튼 라벨 (해당 상태로 전이하는 버튼 텍스트 — 무엇을 하는지 명확히) */
export const RETURN_ACTION_LABEL: Record<ReturnStatus, string> = {
  requested: '수거접수',
  pickup_scheduled: '수거 예약함',
  inbound: '구 제품 입고완료',
  inspected: '검수 완료',
  completed: '전체 완료 처리',
  cancelled: '이 건 취소',
};

/** 현재 상태에서 "지금 무엇을 하면 되는지" 한 줄 안내 */
export const RETURN_STATUS_HINT: Record<ReturnStatus, string> = {
  requested: '구 제품 회수를 접수했습니다. 택배 기사 방문 예약을 잡았으면 「수거 예약함」, 제품이 도착했으면 바로 「구 제품 입고완료」를 누르세요.',
  pickup_scheduled: '수거 예약됨. 구 제품이 매장에 도착하면 「구 제품 입고완료」를 누르세요.',
  inbound: '구 제품이 입고됐습니다(반품창고 확정). 상태 확인이 끝나면 「검수 완료」 또는 바로 「전체 완료 처리」.',
  inspected: '검수 완료. 교환/환불까지 끝났으면 「전체 완료 처리」.',
  completed: '완료된 건입니다.',
  cancelled: '취소된 건입니다.',
};

/** 현재 상태에서 "다음의 대표(primary) 액션" — 이걸 큰 버튼으로. 나머지는 보조 */
export const RETURN_PRIMARY_NEXT: Partial<Record<ReturnStatus, ReturnStatus>> = {
  requested: 'inbound',          // 대개 바로 입고완료(수거예약은 선택)
  pickup_scheduled: 'inbound',
  inbound: 'completed',
  inspected: 'completed',
};

/** 수거방식별 표시 라벨 (목록에서 상태 대신 보여줄 안내) */
export function getReturnDisplayLabel(status: ReturnStatus, pickupMethod?: string | null): string {
  if (status === 'requested') {
    if (pickupMethod === '방문수거') return '수거 예약 필요';
    if (pickupMethod === '택배수거') return '수거 접수됨';
    if (pickupMethod === '직접반납') return '입고 대기';
  }
  return RETURN_STATUS_LABEL[status];
}

/**
 * 반품/교환 구분을 고객 문구용 한글로 (2026-09-17)
 *
 * 알림톡 `return_received` 본문 첫 줄이 "반품·교환 회수 접수가 완료되었습니다" 였다.
 * 사장님 지적: **교환 신청한 고객이 "반품" 글자를 보면 헷갈린다.**
 * returns 에 return_type 이 이미 있는데 알림톡에 안 넘기고 있어서 뭉뚱그렸던 것.
 *
 * 🚨 빈 문자열을 절대 반환하지 않는다 — 알림톡 변수가 비면 문자(SMS)로 대체되고 버튼이 사라진다.
 */
export function returnTypeLabel(returnType?: string | null): string {
  return String(returnType || '').trim() === 'return' ? '반품' : '교환';
}
