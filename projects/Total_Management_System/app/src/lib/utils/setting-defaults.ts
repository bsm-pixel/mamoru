/**
 * 설정 기본값 — 하드코딩 제거용 공유 상수
 * useSetting(key, DEFAULT) 패턴에서 fallback으로 사용
 */

export const DEFAULT_CAT_LABELS: Record<string, string> = {
  BL: '블런트', TH: '씨닝', LO: '롱', SL: '슬라이싱',
  CB: '빗', CS: '케이스', AC: '악세서리', RS: '복원수리',
};

export const DEFAULT_PAYMENT_LABELS: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
};

export const DEFAULT_CHANNEL_LABELS: Record<string, string> = {
  offline: '오프라인', talk: '온라인상담',
};

export const DEFAULT_RFM = {
  recency_vip: 90, recency_dormant: 180, frequency: 3, monetary: 1000000,
};

export const DEFAULT_EXPENSE_CATEGORIES = [
  '택배비', '포장재', '교통비', '사무용품', '식대', '소모품', '임대료', '인건비', '기타',
];
