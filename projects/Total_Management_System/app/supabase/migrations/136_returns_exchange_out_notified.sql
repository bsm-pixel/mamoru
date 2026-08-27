-- 136_returns_exchange_out_notified.sql — 교환 출고 알림톡 발송 여부 (2026-08-27)
-- 판매(offline_sales.shipped_notified_at)와 동일 개념: 교환 출고 송장의 '집하 감지 → 출고 알림톡'을
-- 정확히 1회만 보내기 위한 dedup 플래그. 크론(track-delivery)이 CAS 로 선점하며 채운다.
-- additive only — 기존 데이터/로직 무영향(NULL = 아직 미발송).

ALTER TABLE returns ADD COLUMN IF NOT EXISTS exchange_out_notified_at timestamptz;

COMMENT ON COLUMN returns.exchange_out_notified_at IS '교환 출고 알림톡 발송 시각(집하 감지 시 크론이 채움, NULL=미발송)';
