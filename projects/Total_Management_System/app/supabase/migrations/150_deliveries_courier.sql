-- 150: 납품(deliveries)에 택배사 기록 — 롯데 외 배송·직접수령 경로 지원 (2026-09-15)
--
-- 배경: 롯데 ALPS 송장을 뽑았다가 집하취소하고 우체국으로 직접 보낸 건을 기록할 방법이 없었다.
--   · deliveries 에는 courier_name 컬럼이 아예 없어(28컬럼 실측) 송장번호 = 롯데 송장이라는 암묵 전제
--   · 우체국 송장을 수동 입력하면 크론이 그 번호로 롯데 ALPS 를 매시간 헛조회 → 배송완료 영원히 안 뜸
--   offline_sales / repairs 에는 이미 courier_name 이 있어 3개 도메인 모양을 맞춘다.
--
-- 값 규칙(lib/shipping/couriers.ts 가 SSOT):
--   '롯데택배'  = ALPS 자동추적 대상 (집하감지 → 출고완료, 인수자등록 → 배송완료)
--   그 외       = 자동추적 제외. 출고완료는 수동 처리
--   '직접전달'  = 택배 없이 거래처 직접 수령/방문 전달
--   NULL        = 구데이터 호환 — 롯데로 간주(아래 백필로 대부분 해소)

ALTER TABLE deliveries ADD COLUMN IF NOT EXISTS courier_name text;

COMMENT ON COLUMN deliveries.courier_name IS
  '택배사명. 롯데택배만 ALPS 자동추적 대상. 직접전달=수령. NULL=구데이터(롯데 간주)';

-- 백필: 기존에 송장이 있는 건은 전부 롯데 ALPS 로 뽑은 것이므로 롯데택배로 채운다
UPDATE deliveries
   SET courier_name = '롯데택배'
 WHERE tracking_number IS NOT NULL
   AND courier_name IS NULL;

-- 확인용
SELECT courier_name, count(*) FROM deliveries GROUP BY courier_name ORDER BY 2 DESC;
