import { format, formatDistanceToNow, differenceInCalendarDays } from 'date-fns';
import { ko } from 'date-fns/locale';

/** 원 포맷: 350000 → "350,000원" */
export function formatKRW(amount: number): string {
  return new Intl.NumberFormat('ko-KR').format(amount) + '원';
}

/**
 * 로컬(KST) 날짜를 ISO date string('YYYY-MM-DD')으로 변환
 * `Date.toISOString()`은 UTC 변환을 거치므로 KST 5/1 자정을 4/30으로 잘못 계산하는 버그 회피
 *
 * @example
 *   const today = toLocalDateString(new Date());      // '2026-05-01' (KST 기준)
 *   const monthStart = toLocalDateString(new Date(now.getFullYear(), now.getMonth(), 1));
 */
export function toLocalDateString(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/** 날짜 포맷: "2026-02-16" → "2월 16일" */
export function formatDate(date: string | Date, fmt = 'M월 d일'): string {
  return format(new Date(date), fmt, { locale: ko });
}

/** 날짜 포맷 상세: "2026-02-16T10:00" → "2026.02.16 10:00" */
export function formatDateTime(date: string | Date): string {
  return format(new Date(date), 'yyyy.MM.dd HH:mm', { locale: ko });
}

/** 상대 시간: "3분 전", "2시간 전" */
export function formatRelative(date: string | Date): string {
  return formatDistanceToNow(new Date(date), { addSuffix: true, locale: ko });
}

/** VAT 계산: 총액 → 공급가액 + 부가세(10%) 분리
 * @param vatType 'included'(포함) | 'separate'(별도) | 'none'(미적용) */
export function calcVAT(total: number, vatType: 'included' | 'separate' | 'none' = 'included'): { supply: number; vat: number; payment: number } {
  if (vatType === 'separate') {
    const vat = Math.round(total * 0.1);
    return { supply: total, vat, payment: total + vat };
  }
  if (vatType === 'none') {
    return { supply: total, vat: 0, payment: total };
  }
  // 'included' (기본)
  const supply = Math.round(total / 1.1);
  return { supply, vat: total - supply, payment: total };
}

/** 전화번호 포맷: "01012345678" → "010-1234-5678" */
export function formatPhone(phone: string | null): string {
  if (!phone) return '';
  const num = phone.replace(/\D/g, '');
  if (num.length === 11) return `${num.slice(0, 3)}-${num.slice(3, 7)}-${num.slice(7)}`;
  if (num.length === 10) return `${num.slice(0, 3)}-${num.slice(3, 6)}-${num.slice(6)}`;
  return phone;
}

/** 주문 상태 한글 매핑 */
export const ORDER_STATUS_LABEL: Record<string, string> = {
  pay_wait: '입금대기',
  pay_done: '결제완료',
  preparing: '상품준비중',
  shipping: '배송중',
  delivered: '배송완료',
  confirmed: '구매확정',
  cancel_pending: 'ALPS 취소 필요',
  cancelled: '주문취소',
  refund_request: '환불요청',
  refunded: '환불완료',
};

/** 주문 상태 색상 매핑 (Tailwind 클래스) */
export const ORDER_STATUS_COLOR: Record<string, string> = {
  pay_wait: 'bg-warning-soft text-warning',
  pay_done: 'bg-info-soft text-info',
  preparing: 'bg-info-soft text-info',
  shipping: 'bg-terracotta-soft/30 text-terracotta-deep',
  delivered: 'bg-success-soft text-success',
  confirmed: 'bg-success-soft text-success',
  cancel_pending: 'bg-warning-soft text-warning',
  cancelled: 'bg-error-soft text-error',
  refund_request: 'bg-error-soft text-error',
  refunded: 'bg-neutral-100 text-neutral-500',
};

// ============================================
// Phase 2-1: 상담 상태 라벨/색상
// ============================================

/** 상담 상태 한글 매핑 */
export const CONSULTATION_STATUS_LABEL: Record<string, string> = {
  pending_admin: '대기중',
  suggested: '일정 제안',
  assigned: '딜러 배정',
  confirmed: '확정',
  cancelled: '취소',
  reschedule_requested: '일정 변경 요청',
  change_requested: '일정 변경 요청',
  on_hold: '보류',           // Phase 2-2
  in_progress: '진행중',     // Phase 2-2
  completed: '처리완료',     // Phase 2-2
};

/** 상담 상태 색상 매핑 (Tailwind 클래스) */
export const CONSULTATION_STATUS_COLOR: Record<string, string> = {
  pending_admin: 'bg-warning-soft text-warning',
  suggested: 'bg-info-soft text-info',
  assigned: 'bg-terracotta-soft/30 text-terracotta-deep',
  confirmed: 'bg-success-soft text-success',
  cancelled: 'bg-error-soft text-error',
  reschedule_requested: 'bg-warning-soft text-warning',
  change_requested: 'bg-orange-100 text-orange-700',
  on_hold: 'bg-neutral-100 text-neutral-500',    // Phase 2-2
  in_progress: 'bg-info-soft text-info',          // Phase 2-2
  completed: 'bg-success-soft text-success',      // Phase 2-2
};

// ============================================
// D-day / 날짜 그룹 유틸
// ============================================

/** D-day 라벨: "오늘", "내일", "D-3", "D+2(지남)" */
export function formatDday(dateStr: string | null): { label: string; isPast: boolean; isToday: boolean } {
  if (!dateStr) return { label: '', isPast: false, isToday: false };
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(dateStr + 'T00:00:00');
  const diff = differenceInCalendarDays(target, today);

  if (diff === 0) return { label: '오늘', isPast: false, isToday: true };
  if (diff === 1) return { label: '내일', isPast: false, isToday: false };
  if (diff === -1) return { label: '어제', isPast: true, isToday: false };
  if (diff > 1) return { label: `D-${diff}`, isPast: false, isToday: false };
  return { label: `D+${Math.abs(diff)}(지남)`, isPast: true, isToday: false };
}

/** 날짜 그룹 헤더: "2/17 (월) · 오늘" */
export function formatDateGroup(dateStr: string): string {
  const d = new Date(dateStr + 'T00:00:00');
  const dayNames = ['일', '월', '화', '수', '목', '금', '토'];
  const base = `${d.getMonth() + 1}/${d.getDate()} (${dayNames[d.getDay()]})`;
  const { label } = formatDday(dateStr);
  return label ? `${base} · ${label}` : base;
}

/** 상담 타입 한글 매핑 */
export const CONSULTATION_TYPE_LABEL: Record<string, string> = {
  store_visit: '매장 방문',
  field_request: '출장 요청',
  talk_consult: '온라인상담',    // Phase 2-2
};

// ============================================
// Phase 7: 복원수리 상태 라벨/색상
// ============================================

/** 복원수리 상태 한글 매핑 */
export const REPAIR_STATUS_LABEL: Record<string, string> = {
  intake: '접수',
  pickup_scheduled: '입고대기중',
  cost_notified: '진행중',
  payment_confirmed: '진행중',
  repairing: '진행중',
  ready_to_ship: '출고대기',  // R1: 활성 상태로 승격
  shipped: '출고 완료',
  delivered: '배송완료',
  completed: '완료',
  cancelled: '취소',
  // 레거시
  picked_up: '입고완료',
  inspecting: '검수중',
};

/** 복원수리 상태 색상 매핑 (Tailwind 클래스) */
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
