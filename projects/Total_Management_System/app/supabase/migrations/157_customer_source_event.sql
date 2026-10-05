-- 157: 고객 출처(customer_source)에 'event' · 'stock_sale' 추가 (2026-10-05)
--
-- 사고: 이벤트 접수로 처음 온 고객(박지은, EV-20261004-001)이 고객 등록에 실패 → 판매가 "고객 없음"으로 저장
--       → [택배 발송(송장 생성)]이 "고객 정보가 없어 송장 생성 불가"로 막힘.
-- 원인: customer_source enum = ('imweb','consultation','as','manual') 뿐인데
--       코드(event 접수·재고판매 접수·판매 전환)는 source='event' / 'stock_sale' 로 INSERT
--       → 22P02 invalid input value for enum → 실패를 로그만 남기고 진행(조용한 누락).
--       기존 이벤트 구매자는 전부 "이미 있던 고객"이라 매칭만 되고 INSERT 가 안 일어나 그동안 안 드러났다.
--
-- 이 마이그레이션은 값만 추가한다(기존 행·기본값 영향 없음). 실행 전에도 코드가 'manual' 로 재시도하므로 순서 무관.
-- ⚠️ ALTER TYPE ... ADD VALUE 는 한 줄씩 따로 실행돼야 한다(같은 트랜잭션에서 새 값을 바로 쓰면 오류) — 아래 그대로 실행하면 된다.

ALTER TYPE customer_source ADD VALUE IF NOT EXISTS 'event';
ALTER TYPE customer_source ADD VALUE IF NOT EXISTS 'stock_sale';

-- 확인용 (event, stock_sale 이 목록에 보이면 성공)
SELECT enumlabel FROM pg_enum
WHERE enumtypid = 'customer_source'::regtype
ORDER BY enumsortorder;
