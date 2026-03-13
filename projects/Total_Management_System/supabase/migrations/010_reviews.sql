-- ============================================================
-- 010_reviews.sql — 통합 리뷰 시스템 테이블
-- Supabase SQL Editor에서 실행
-- ============================================================

CREATE TABLE reviews (
  id           UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  review_id    TEXT NOT NULL UNIQUE,                                        -- RV-YYYYMMDD-NNN
  created_at   TIMESTAMPTZ DEFAULT now(),
  type         TEXT NOT NULL CHECK (type IN ('consult','repair','purchase')),
  subtype      TEXT,                                                        -- store_visit, field_request, talk_consult, restoration 등
  name         TEXT NOT NULL,
  phone        TEXT,
  stars        SMALLINT NOT NULL CHECK (stars BETWEEN 1 AND 5),
  content      TEXT NOT NULL,
  photo_urls   TEXT[] DEFAULT '{}',                                         -- Supabase Storage URL 배열
  source_id    TEXT,                                                        -- consultation.unique_id 또는 repair.as_id
  status       TEXT DEFAULT 'pending' CHECK (status IN ('pending','approved','hidden')),
  approved_at  TIMESTAMPTZ,
  product      TEXT,
  meta         JSONB DEFAULT '{}'                                           -- 부가정보 (날짜, 시간, 상담유형 등)
);

CREATE INDEX idx_reviews_status ON reviews(status);
CREATE INDEX idx_reviews_type ON reviews(type);
CREATE INDEX idx_reviews_source ON reviews(source_id);

-- ── RLS ──
ALTER TABLE reviews ENABLE ROW LEVEL SECURITY;

-- 공개: approved 리뷰만 조회 (anon 포함)
CREATE POLICY "approved_reviews_public" ON reviews
  FOR SELECT USING (status = 'approved');

-- 인증 사용자: 전체 CRUD (TMS 관리용)
CREATE POLICY "admin_full_access" ON reviews
  FOR ALL USING (auth.role() = 'authenticated');

-- anon INSERT 허용 (고객 리뷰 제출용)
CREATE POLICY "anon_insert" ON reviews
  FOR INSERT WITH CHECK (true);
