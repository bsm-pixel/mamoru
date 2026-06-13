-- 102: 고객 활동명(매장 사용 이름) + 직급 (2026-06-12)
-- 출장 방문 시 실명(김순실)으로 못 알아보는 경우 多 → 매장 활동명(하은)·직급(원장)으로 식별·호칭.
-- 매장명은 기존 customers.company_name 재사용. 활동명/직급은 신설.
-- consultations 에도 denormalize(상담 카드 표시용, 기존 name/address 패턴과 동일).

ALTER TABLE customers ADD COLUMN IF NOT EXISTS activity_name text;   -- 활동명 (예: 하은)
ALTER TABLE customers ADD COLUMN IF NOT EXISTS position text;        -- 직급 (예: 원장/디자이너)

ALTER TABLE consultations ADD COLUMN IF NOT EXISTS activity_name text;
ALTER TABLE consultations ADD COLUMN IF NOT EXISTS position text;

COMMENT ON COLUMN customers.activity_name IS '매장 사용 이름(활동명). 실명과 다를 때 호칭용';
COMMENT ON COLUMN customers.position IS '직급(원장/디자이너 등)';
