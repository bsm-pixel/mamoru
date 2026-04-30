-- 070: offline_sales.source_consultation_id (출장/매장상담 → 판매 link)
-- consultation 후 그 거래가 판매로 이어진 경우 명시적 연결
-- 리뷰 약속/발송이 원본 상담 한 곳에서만 관리되도록 mirror 모드 트리거
-- Supabase SQL Editor에서 실행

-- 1) FK 컬럼 추가
ALTER TABLE offline_sales
  ADD COLUMN IF NOT EXISTS source_consultation_id UUID NULL REFERENCES consultations(id);

COMMENT ON COLUMN offline_sales.source_consultation_id
  IS '출처 상담 ID. 있으면 리뷰 카드는 mirror 모드(원본 상담에서 관리). 수동 판매/워크인은 NULL.';

-- 2) 조회 성능용 인덱스 (필터링 + 역방향 JOIN용)
CREATE INDEX IF NOT EXISTS idx_offline_sales_source_consultation
  ON offline_sales (source_consultation_id)
  WHERE source_consultation_id IS NOT NULL;
