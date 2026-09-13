/**
 * 상담 리뷰 subtype 어휘 SSOT (서버 공용 — 순수 함수).
 *
 * 배경(2026-09-13): 판매건(OS-) 상담 후기는 리뷰 약속 칩이 없으면 판매채널(sale_channel) 원시값으로
 * subtype 을 채웠다. sale_channel 4분류(store/field/online/talk)는 리뷰 어휘(store_visit/field_request/talk_consult)와
 * 달라서 'store' 가 그대로 저장 → 관리자 "상담·매장" / 고객 페이지 "상담"(라벨 없음)으로 갈라져 표시됐다.
 * → 저장 시점에 리뷰 어휘로 정규화하고, 작성 폼 라벨도 고객 노출 칩과 같은 표기로 통일한다.
 */

/** 판매채널 → 상담 리뷰 subtype. 대응값 없으면(offline/online 등 레거시) 원값 유지 — 기존 동작 보존 */
export function saleChannelToConsultSubtype(channel: string | null | undefined): string {
  if (channel === 'store') return 'store_visit';
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
