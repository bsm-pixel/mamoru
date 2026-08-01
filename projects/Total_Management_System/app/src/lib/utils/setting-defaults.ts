/**
 * 설정 기본값 — 하드코딩 제거용 공유 상수
 * useSetting(key, DEFAULT) 패턴에서 fallback으로 사용
 */

export const DEFAULT_CAT_LABELS: Record<string, string> = {
  BL: '블런트', TH: '씨닝', LO: '롱', SL: '슬라이싱',
  CB: '빗', CS: '케이스', AC: '악세서리', RS: '복원수리',
  EVENT: 'EVENT', // 재고 전환 이벤트 품목 (category='EVENT')
  LS: '재고판매', // 재고판매 카탈로그 품목 (category='LS') — 아임웹 커스텀 카탈로그로 노출
};

/**
 * 시스템 고정 카테고리 — 코드가 직접 연동(EVENT 허브가 category='EVENT'로 필터).
 * 설정 목록에서 삭제돼도 항상 보장되고, 설정 UI에서 삭제 불가.
 */
export const SYSTEM_CATEGORIES = ['EVENT'];

/**
 * TMS 화면에서 숨기는 카테고리 — 코드/데이터는 남겨두되(휴면) 드롭다운·필터·설정 목록에 노출하지 않는다.
 * LS(재고판매): 커스텀 판매 중단, 온라인 판매는 아임웹 아울렛으로 이관(2026-08-01). 라벨 매핑은 유지(잔존 데이터 표시용).
 */
export const HIDDEN_CATEGORIES = ['LS'];

export const DEFAULT_PAYMENT_LABELS: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
};

export const DEFAULT_CHANNEL_LABELS: Record<string, string> = {
  store: '매장', field: '출장', talk: '톡', online: '온라인(아임웹)', offline: '오프라인',
};

export const DEFAULT_RFM = {
  recency_vip: 90, recency_dormant: 180, frequency: 3, monetary: 1000000,
};

export const DEFAULT_EXPENSE_CATEGORIES = [
  '택배비', '포장재', '교통비', '사무용품', '식대', '소모품', '임대료', '인건비', '기타',
];

/** 현금흐름(cashflow) 입금 카테고리 기본값 */
export const DEFAULT_CASHFLOW_INCOME_CATEGORIES = [
  '매출입금', '기타입금', '환불수령', '투자금',
];

/** 현금흐름(cashflow) 출금 카테고리 기본값 */
export const DEFAULT_CASHFLOW_EXPENSE_CATEGORIES = [
  '매입결제', '경비', '세금', '인건비', '임대료', '기타출금',
];
