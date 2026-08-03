-- 122: 게스트 판매(customer_id NULL)를 전화번호로 고객에 연결 (백필)
--
-- 배경: EVENT/재고판매 전환 또는 수동등록에서 접수 시 customer_id 가 안 잡히면
--   판매가 '게스트'(customer_id NULL)로 저장 → 송장 생성(customer_id 필수, ship/route.ts:31)·
--   미수금 집계 등이 막힌다. (예: OS-20260801-001 김용준 — 송장 생성 불가)
-- 코드는 convert-to-sale.ts 에서 재발 방지(전환 시 전화/이름으로 매칭·생성) 완료(2026-08-02).
-- 이 마이그는 '이미 게스트로 저장된 과거 건'을 전화번호로 소급 연결한다.
--
-- 매칭 규칙: 판매 전화(digits) = customers.phone_normalized.
--   동일 전화에 고객이 여럿이면 '가장 오래된 고객'(created_at ASC) — matchOrCreateCustomer 와 동일한 deterministic 규칙.
--   병합/활동명 등 정리는 [merge_customers] 로 사후 처리 가능.

-- ────────────────────────────────────────────────────────────────
-- [1] 진단 (읽기전용) — 무엇이 어느 고객에 연결될지 먼저 확인. 이 SELECT 만 돌려보고 결과 검토 후 [2] 실행.
-- ────────────────────────────────────────────────────────────────
-- SELECT os.sale_number, os.sale_date, os.customer_name, os.customer_phone,
--        oldest.id AS will_link_customer_id
-- FROM offline_sales os
-- JOIN (
--   SELECT DISTINCT ON (phone_normalized) id, phone_normalized
--   FROM customers
--   WHERE phone_normalized IS NOT NULL AND phone_normalized <> ''
--   ORDER BY phone_normalized, created_at ASC
-- ) oldest
--   ON regexp_replace(COALESCE(os.customer_phone, ''), '\D', '', 'g') = oldest.phone_normalized
-- WHERE os.customer_id IS NULL
--   AND COALESCE(os.customer_phone, '') <> ''
-- ORDER BY os.sale_date DESC;

-- ────────────────────────────────────────────────────────────────
-- [2] 보정 (UPDATE) — 진단 결과가 맞으면 실행.
-- ────────────────────────────────────────────────────────────────
UPDATE offline_sales os
SET customer_id = oldest.id
FROM (
  SELECT DISTINCT ON (phone_normalized) id, phone_normalized
  FROM customers
  WHERE phone_normalized IS NOT NULL AND phone_normalized <> ''
  ORDER BY phone_normalized, created_at ASC
) oldest
WHERE os.customer_id IS NULL
  AND COALESCE(os.customer_phone, '') <> ''
  AND regexp_replace(COALESCE(os.customer_phone, ''), '\D', '', 'g') = oldest.phone_normalized;

-- [참고] 연결된 판매 중 미결제(unpaid/partial) 건이 있으면, 해당 고객의 미수금은
--   다음 금액변동 시 recalcOutstanding 로 자동 정정됨(또는 고객패널에서 재계산). EVENT/재고 전환분은 결제완료라 무영향.
