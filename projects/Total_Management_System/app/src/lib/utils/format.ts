import { format, formatDistanceToNow, differenceInCalendarDays } from 'date-fns';
import { ko } from 'date-fns/locale';
import { todayKST, addDays } from '@/lib/utils/kst';
import { toneClasses } from '@/lib/ui/status-tone';

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

/* ── 날짜 표기 (2026-11-15 통일) ─────────────────────────────────────────────
 * 전엔 11월 15일 / 2026.11.15 / 11/15 / 2026-11-14 / 9/15 다섯 가지가 섞여 있었다. 화면에 날짜를 보일 때는 아래 넷만 쓴다.
 *   목록   formatDate(d)           → 11.15 (일)        · 오늘 / 어제 / 내일은 글자로 · 올해가 아니면 2025.11.15 (토)
 *   목록+시각 formatDateTimeShort(d) → 11.15 (일) 14:30  · 오늘 14:30
 *   상세   formatDateFull(d)       → 2026.11.15 (일)
 *   상세+시각 formatDateTime(d)     → 2026.11.15 (일) 14:30
 * 모두 한국 시간 기준(브라우저·서버 시간대와 무관). 날짜가 비었거나 잘못됐으면 빈 글자.
 * 달력 계산 · 서버로 보내는 값은 여기 말고 lib/utils/kst.ts 를 쓴다.
 */
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const KST_MS = 9 * 60 * 60 * 1000;

/** 어떤 값이든 한국 날짜('YYYY-MM-DD')와 시각('HH:mm')으로. 날짜만 있는 글자는 그대로 쓴다(시간대 변환 없음) */
function kstParts(input: string | Date | number | null | undefined): { key: string; hm: string } | null {
  if (input === null || input === undefined || input === '') return null;
  if (typeof input === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(input)) return { key: input, hm: '' };
  const t = input instanceof Date ? input.getTime() : new Date(input).getTime();
  if (!Number.isFinite(t)) return null;
  const iso = new Date(t + KST_MS).toISOString();
  return { key: iso.slice(0, 10), hm: iso.slice(11, 16) };
}

function dayLabel(key: string, full: boolean): string {
  const today = todayKST();
  if (!full) {
    if (key === today) return '오늘';
    if (key === addDays(today, -1)) return '어제';
    if (key === addDays(today, 1)) return '내일';
  }
  const [y, m, d] = key.split('-').map(Number);
  const dow = DOW[new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
  const md = `${key.slice(5, 7)}.${key.slice(8, 10)} (${dow})`;
  return full || key.slice(0, 4) !== today.slice(0, 4) ? `${y}.${md}` : md;
}

/** 목록용 날짜: "11.15 (일)" · 오늘/어제/내일은 글자로.
 *  fmt 를 주면 예전처럼 그 모양 그대로(시각만 'HH:mm', 기간 표시 'M.d' 처럼 특수한 곳 전용) */
export function formatDate(date: string | Date | number | null | undefined, fmt?: string): string {
  if (fmt) {
    if (date === null || date === undefined || date === '') return '';
    const d = new Date(date);
    return Number.isFinite(d.getTime()) ? format(d, fmt, { locale: ko }) : '';
  }
  const p = kstParts(date);
  return p ? dayLabel(p.key, false) : '';
}

/** 상세용 날짜: "2026.11.15 (일)" */
export function formatDateFull(date: string | Date | number | null | undefined): string {
  const p = kstParts(date);
  return p ? dayLabel(p.key, true) : '';
}

/** 목록용 날짜+시각: "11.15 (일) 14:30" · "오늘 14:30" */
export function formatDateTimeShort(date: string | Date | number | null | undefined): string {
  const p = kstParts(date);
  return p ? `${dayLabel(p.key, false)}${p.hm ? ' ' + p.hm : ''}` : '';
}

/** 상세용 날짜+시각: "2026.11.15 (일) 14:30" */
export function formatDateTime(date: string | Date | number | null | undefined): string {
  const p = kstParts(date);
  return p ? `${dayLabel(p.key, true)}${p.hm ? ' ' + p.hm : ''}` : '';
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
export function formatPhone(phone: string | null | undefined): string {
  if (!phone) return '';
  // 2026-11-15 보여 줄 때는 항상 하이픈 — 서울(02) · 대표번호(1588) · 안심번호(12자리)까지
  if (phone.trim().startsWith('+')) return phone; // 해외 번호는 그대로
  const num = phone.replace(/\D/g, '');
  if (num.startsWith('02') && (num.length === 9 || num.length === 10)) return `02-${num.slice(2, num.length - 4)}-${num.slice(-4)}`;
  if (num.length === 12) return `${num.slice(0, 4)}-${num.slice(4, 8)}-${num.slice(8)}`;
  if (num.length === 11) return `${num.slice(0, 3)}-${num.slice(3, 7)}-${num.slice(7)}`;
  if (num.length === 10) return `${num.slice(0, 3)}-${num.slice(3, 6)}-${num.slice(6)}`;
  if (num.length === 8 && num.startsWith('1')) return `${num.slice(0, 4)}-${num.slice(4)}`;
  return phone;
}

/** 주문 상태 한글 매핑 */
export const ORDER_STATUS_LABEL: Record<string, string> = {
  pay_wait: '입금대기',
  pay_done: '결제완료',
  preparing: '상품준비중',
  ready_to_ship: '배송대기',
  shipping: '배송중',
  delivered: '배송완료',
  confirmed: '구매확정',
  cancel_pending: 'ALPS 취소 필요',
  cancelled: '주문취소',
  refund_request: '환불요청',
  refunded: '환불완료',
};

/** 주문 상태 색상 매핑 (Tailwind 클래스) */
export const ORDER_STATUS_COLOR: Record<string, string> = toneClasses({
  pay_wait: 'wait', pay_done: 'progress', preparing: 'progress', ready_to_ship: 'wait', shipping: 'progress', delivered: 'done', confirmed: 'done', cancel_pending: 'wait', cancelled: 'cancel', refund_request: 'problem', refunded: 'cancel',
}); // 2026-11-15 상태 색 표(lib/ui/status-tone.ts)

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
export const CONSULTATION_STATUS_COLOR: Record<string, string> = toneClasses({
  pending_admin: 'wait', suggested: 'wait', assigned: 'progress', confirmed: 'progress', cancelled: 'cancel', reschedule_requested: 'wait', change_requested: 'wait', on_hold: 'cancel', in_progress: 'progress', completed: 'done',
}); // 2026-11-15 상태 색 표(lib/ui/status-tone.ts)

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

/**
 * 판매 채널(offline_sales.sale_channel) 한글 매핑 (2026-07-17)
 * 신규 입력: store(매장)/field(출장)/talk(톡)/online(온라인·아임웹).
 * 레거시: offline(오프라인)은 소급 구분 불가라 그대로 표시.
 */
export const SALE_CHANNEL_LABEL: Record<string, string> = {
  store: '매장',
  field: '출장',
  talk: '톡',
  online: '온라인(아임웹)',
  offline: '오프라인', // 레거시 (구 데이터)
};

/** 고객유형 라벨/색 — SSOT (기존 화면마다 재정의되며 딜러 색이 파랑↔보라로 충돌하던 것 통일) */
export const CUSTOMER_TYPE_LABEL: Record<string, string> = {
  retail: '일반',
  online: '온라인',
  dealer: '딜러',
  academy: '아카데미',
  supplier: '매입처',
};
export const CUSTOMER_TYPE_COLOR: Record<string, string> = {
  retail: 'bg-neutral-100 text-neutral-600',
  online: 'bg-blue-100 text-blue-700',
  dealer: 'bg-purple-100 text-purple-700',
  academy: 'bg-emerald-100 text-emerald-700',
  supplier: 'bg-amber-100 text-amber-700',
};

/** 상담 유형 → 판매 채널 자동 매핑 (상담 연결 판매 채널 자동설정용). 없으면 '' */
export function channelFromConsultationType(t?: string | null): string {
  if (t === 'store_visit') return 'store';
  if (t === 'field_request') return 'field';
  if (t === 'talk_consult') return 'talk';
  return '';
}

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
  ready_to_ship: '작업완료',  // 2026-10-19 출고대기→작업완료 (표시만)
  shipped: '출고 완료',
  delivered: '배송완료',
  completed: '완료',
  cancelled: '취소',
  // 레거시
  picked_up: '입고완료',
  inspecting: '검수중',
};

/** 복원수리 상태 색상 매핑 (Tailwind 클래스) */
export const REPAIR_STATUS_COLOR: Record<string, string> = toneClasses({
  intake: 'progress', pickup_scheduled: 'progress', cost_notified: 'wait', payment_confirmed: 'progress', repairing: 'progress', shipped: 'done', delivered: 'done', completed: 'done', cancelled: 'cancel', picked_up: 'progress', inspecting: 'progress', ready_to_ship: 'wait',
}); // 2026-11-15 상태 색 표(lib/ui/status-tone.ts)
