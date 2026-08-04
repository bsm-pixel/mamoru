-- 125_repairs_visit_remind.sql
-- 직접방문(매장방문) 리마인드 알림톡 중복방지 플래그
-- 크론 api/cron/repair-visit-remind 가 24h/2h 발송 전 선(先)마킹에 사용
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS visit_remind_24h_sent_at timestamptz;
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS visit_remind_2h_sent_at  timestamptz;
