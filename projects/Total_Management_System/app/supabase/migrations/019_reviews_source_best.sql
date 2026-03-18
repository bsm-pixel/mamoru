-- 019: reviews 테이블에 source, is_best 컬럼 추가
-- source: 리뷰 출처 (site=자체수집, naver=네이버플레이스)
-- is_best: 베스트 리뷰 여부

ALTER TABLE reviews ADD COLUMN IF NOT EXISTS source text NOT NULL DEFAULT 'site';
ALTER TABLE reviews ADD COLUMN IF NOT EXISTS is_best boolean NOT NULL DEFAULT false;

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_reviews_source ON reviews (source);
CREATE INDEX IF NOT EXISTS idx_reviews_is_best ON reviews (is_best) WHERE is_best = true;

COMMENT ON COLUMN reviews.source IS '리뷰 출처: site(자체), naver(네이버플레이스)';
COMMENT ON COLUMN reviews.is_best IS '베스트 리뷰 여부';
