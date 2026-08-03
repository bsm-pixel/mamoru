-- 124_review_event_entry_start.sql
-- 리뷰 이벤트 응모 시작일(선택) — 비우면 그 달 1일부터, 값 있으면 그 날부터 응모자 집계
-- 용도: 첫 회차 등 과거 후기까지 포함해 선정하고 싶을 때 시작일을 앞당겨 지정
ALTER TABLE review_event_config ADD COLUMN IF NOT EXISTS entry_start timestamptz;
