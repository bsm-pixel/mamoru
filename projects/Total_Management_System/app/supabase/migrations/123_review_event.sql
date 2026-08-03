-- 123_review_event.sql
-- 리뷰 이벤트(월별 베스트/등수) 관리 — 사장님 2026-08-03 Supabase 수동 실행 완료분의 리포 기록본
-- 목적: 월별 이벤트 설정(상품/마감일/발표상태) + reviews 에 당첨 등수 마킹(별도테이블 아닌 SSOT)

-- 월별 이벤트 설정
CREATE TABLE IF NOT EXISTS review_event_config (
  month text PRIMARY KEY,                 -- 'YYMM' (예: '2508')
  deadline timestamptz,                   -- 응모 마감일(히어로 카운트다운)
  announce_at timestamptz,                -- 발표 예정일(안내용)
  hero_image_url text,                    -- 히어로 1등 상품 이미지(선택)
  prizes jsonb NOT NULL DEFAULT '[]',     -- [{rank,name,desc,image_url,count}]
  status text NOT NULL DEFAULT 'draft',   -- draft(편집중) / live(진행중) / announced(발표됨)
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

-- 당첨자 = reviews 레코드에 마킹 (별도 테이블 X → SSOT)
ALTER TABLE reviews ADD COLUMN IF NOT EXISTS event_month text;          -- 어느 달 당첨 'YYMM'
ALTER TABLE reviews ADD COLUMN IF NOT EXISTS event_rank int;            -- 1/2/3, null=미당첨
ALTER TABLE reviews ADD COLUMN IF NOT EXISTS event_display_name text;   -- 마스킹 표시명 오버라이드(선택)
ALTER TABLE reviews ADD COLUMN IF NOT EXISTS event_route text;          -- '매장 구매 · 블런트' 경로 라벨(선택)
CREATE INDEX IF NOT EXISTS idx_reviews_event ON reviews(event_month, event_rank);
