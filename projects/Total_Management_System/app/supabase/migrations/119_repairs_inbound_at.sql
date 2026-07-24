-- 119: 복원수리 입고일(inbound_at) — 2026-07-24
--
-- 배경: 목록에 "입고일"을 보여주고 싶은데, confirmed_at 은 '신규접수 → 접수확인' 클릭 시점이라
--   실제 물리 입고 시점과 다르다. 사장님이 "입고확인 & 비용안내 발송"을 누르는 순간이
--   가위가 손에 들어와 검수·비용을 확정하는 시점 → 그때를 inbound_at 으로 기록한다.
--
-- 세팅: 비용안내(as_cost_notice) 발송 = status cost_notified 전이 시 inbound_at 을 처음 한 번 찍는다.
--   (이미 값이 있으면 덮어쓰지 않음 — 재발송해도 최초 입고시점 유지)
--
-- 되돌리기: ALTER TABLE repairs DROP COLUMN inbound_at;

ALTER TABLE repairs
  ADD COLUMN IF NOT EXISTS inbound_at timestamptz;

COMMENT ON COLUMN repairs.inbound_at
  IS '입고일 — 입고확인&비용안내(as_cost_notice) 발송 시점. 목록 "입고일" 컬럼용 (119, 2026-07-24)';
