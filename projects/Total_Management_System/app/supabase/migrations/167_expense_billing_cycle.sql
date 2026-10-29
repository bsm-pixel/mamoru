-- 167: 고정 경비 「연간」 구분 (2026-10-29)
--
-- 배경(사장님): Make 연간 구독($54, 2026-10-01~2027-10-01)처럼 1년에 한 번 내는 고정비가 결제한 달에만 크게 잡혀
--   "한 달 고정비"가 들쭉날쭉해 보인다.
-- 원칙: 장부(경비·리포트 손익)는 현금주의 그대로 — 결제일에 전액. billing_cycle 은 경비 화면의 「월 환산 고정비」(연간 ÷ 12)
--   계산에만 쓴다. kind='fixed' 일 때만 의미가 있고, NULL = 매달(기존 데이터 전부).

ALTER TABLE expenses             ADD COLUMN IF NOT EXISTS billing_cycle text CHECK (billing_cycle IN ('monthly', 'annual'));
ALTER TABLE expense_inbox        ADD COLUMN IF NOT EXISTS billing_cycle text CHECK (billing_cycle IN ('monthly', 'annual'));
ALTER TABLE expense_vendor_rules ADD COLUMN IF NOT EXISTS billing_cycle text CHECK (billing_cycle IN ('monthly', 'annual'));

-- 이미 확인 대기에 있는 Make 연간 구독 (판독 규칙은 이 마이그 이후 들어오는 메일부터 자동 표시)
UPDATE expense_inbox SET billing_cycle = 'annual'
 WHERE status = 'pending' AND vendor = 'Make' AND subject ILIKE '%Celonis%';

-- 확인용
SELECT table_name, column_name, data_type FROM information_schema.columns
 WHERE column_name = 'billing_cycle' ORDER BY table_name;
