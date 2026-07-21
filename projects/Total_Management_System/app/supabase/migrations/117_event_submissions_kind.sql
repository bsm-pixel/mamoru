-- 117: event_submissions 에 kind 판별자 추가 — EVENT / 재고판매(LS) 분리 (2026-07-21)
--
-- 배경: 재고판매(LS)는 EVENT 와 접수→입금→판매전환 흐름이 동일해 event_submissions 를 재사용한다.
--   단, 어드민 화면은 분리(EVENT 메뉴 / 재고판매 메뉴)해야 하므로 행을 구분할 판별자가 필요.
--
--   kind='event'       기존 EVENT 접수 (기본값)
--   kind='stock_sale'  재고판매 접수 (신규)
--
-- ⚠️ 재고/수량 로직 무관 — 접수 분류 컬럼만 추가.
--
-- 되돌리기: ALTER TABLE event_submissions DROP COLUMN kind;

ALTER TABLE event_submissions
  ADD COLUMN IF NOT EXISTS kind text NOT NULL DEFAULT 'event';

COMMENT ON COLUMN event_submissions.kind
  IS '접수 종류: event(이벤트) | stock_sale(재고판매). 어드민 화면 분리용 (117, 2026-07-21)';

-- 기존 접수는 전부 EVENT 였다 → 명시적으로 'event' 백필 (기본값이 있어도 안전하게)
UPDATE event_submissions SET kind = 'event' WHERE kind IS NULL OR kind = '';

CREATE INDEX IF NOT EXISTS idx_event_submissions_kind ON event_submissions(kind, status, created_at DESC);
