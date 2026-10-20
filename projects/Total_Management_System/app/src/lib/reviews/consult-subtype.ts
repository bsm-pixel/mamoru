/**
 * 상담 리뷰 subtype 어휘 SSOT (서버 공용 — 순수 함수).
 *
 * 배경(2026-09-13): 판매건(OS-) 상담 후기는 리뷰 약속 칩이 없으면 판매채널(sale_channel) 원시값으로
 * subtype 을 채웠다. sale_channel 4분류(store/field/online/talk)는 리뷰 어휘(store_visit/field_request/talk_consult)와
 * 달라서 'store' 가 그대로 저장 → 관리자 "상담·매장" / 고객 페이지 "상담"(라벨 없음)으로 갈라져 표시됐다.
 * → 저장 시점에 리뷰 어휘로 정규화하고, 작성 폼 라벨도 고객 노출 칩과 같은 표기로 통일한다.
 */

/** 판매채널 → 상담 리뷰 subtype. 레거시 'offline'(4분류 이전 대면판매)=직접방문(사장님 결정 2026-09-13). 대응값 없으면(online 등) 원값 유지 */
export function saleChannelToConsultSubtype(channel: string | null | undefined): string {
  if (channel === 'store' || channel === 'offline') return 'store_visit';
  if (channel === 'field') return 'field_request';
  if (channel === 'talk') return 'talk_consult';
  return channel || '';
}

/** 상담 subtype 라벨 — 고객 노출 칩(page_reviews 등 SUBTYPE_LABELS)과 동일 표기 */
export const CONSULT_SUBTYPE_LABELS: Record<string, string> = {
  store_visit: '직접방문',
  field_request: '출장',
  talk_consult: '톡상담',
};

/** 작성 폼 헤더 라벨: '상담 · 직접방문' (라벨 없으면 '상담') */
export function consultTypeLabel(subtype: string | null | undefined): string {
  const sub = subtype ? CONSULT_SUBTYPE_LABELS[subtype] : '';
  return sub ? `상담 · ${sub}` : '상담';
}

/**
 * 리뷰 유형별 세부유형(subtype) 허용값 SSOT — 리뷰관리 수정 모달 칩 + PATCH 검증 공용 (2026-10-20).
 * 배경: 흐름을 잘못 골라 보낸 후기(제품구매→상담 등)를 고칠 때 subtype 선택 칸이 없어 null 로 남고,
 * 제품명도 그대로 남아 고객 후기벽에 상담 후기인데 제품명 칩이 노출됐다.
 * value '' = subtype 없음(null 저장). 첫 항목 = 유형 변경 시 기본값.
 */
export type ReviewType = 'consult' | 'repair' | 'purchase';

export const REVIEW_SUBTYPE_OPTIONS: Record<ReviewType, { value: string; label: string }[]> = {
  consult: [
    { value: 'store_visit', label: '직접방문' },
    { value: 'field_request', label: '출장' },
    { value: 'talk_consult', label: '톡상담' },
  ],
  repair: [
    { value: 'restoration', label: '기본' },
    { value: 'direct_visit', label: '직접방문' },
    { value: 'parcel_pickup', label: '방문수거' },
    { value: 'self_ship', label: '직접발송' },
  ],
  purchase: [
    { value: '', label: '일반' },
    { value: 'event', label: '이벤트참여' },
  ],
};

export function isReviewType(t: unknown): t is ReviewType {
  return t === 'consult' || t === 'repair' || t === 'purchase';
}

/** 해당 유형에서 허용되는 subtype 인지 (null/'' = 값 없음) */
export function isAllowedReviewSubtype(type: ReviewType, subtype: string | null | undefined): boolean {
  return REVIEW_SUBTYPE_OPTIONS[type].some(o => o.value === (subtype || ''));
}
