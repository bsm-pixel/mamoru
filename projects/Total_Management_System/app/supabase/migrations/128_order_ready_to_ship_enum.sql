-- 128: 주문(orders) 배송대기 상태(ready_to_ship) 추가 — 아임웹 배송대기 미러
--
-- 배경(2026-08-23): 아임웹 주문 배송상태를 API로 역동기 가능함을 실측 확인.
--   place=배송대기 / invoice=송장등록(배송대기 유지) / send=배송중.
--   그동안 주문은 송장생성 시 곧바로 'shipping'(배송중)으로 점프했으나,
--   송장 발급 ≠ 출고다(복원수리·납품엔 이미 적용된 원칙). 주문도 3단계로 통일:
--   송장생성 = ready_to_ship(배송대기) → 집하 = shipping(배송중) → delivered(배송완료)
--
-- ⚠️ ALTER TYPE ... ADD VALUE 는 트랜잭션 블록 안에서 실행 불가 +
--    같은 트랜잭션에서 새 값 사용 불가. 이 파일(128)을 단독 실행(커밋)한 뒤
--    129(get_order_counts 갱신)를 실행할 것. Supabase SQL Editor에서 128 → 129 순서.

ALTER TYPE order_status ADD VALUE IF NOT EXISTS 'ready_to_ship';
