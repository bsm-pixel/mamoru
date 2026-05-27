-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 094_review_promised_type — 리뷰 약속 유형 컬럼 추가                ║
-- ║                                                                    ║
-- ║ 배경 (2026-05-27 사장님 신고):                                     ║
-- ║   OS-20260525-004 자동 후기요청 발송 실패 — track-delivery cron 이 ║
-- ║   매출 유형 무관하게 reviewType='purchase' 고정 발송 → 온라인상담  ║
-- ║   매출에 잘못된 템플릿 적용. 약속 토글이 유형 정보 미저장 구조라   ║
-- ║   자동 발송 시 매출 종류 판별 불가.                                ║
-- ║                                                                    ║
-- ║ 개선:                                                              ║
-- ║   3개 테이블(offline_sales/repairs/consultations)에                ║
-- ║   review_promised_type 컬럼 추가. 사장님이 약속 토글 ON 시         ║
-- ║   [복원수리/상담/제품구매] 유형을 명시적으로 선택 → 자동 발송 시   ║
-- ║   해당 유형의 솔라피 템플릿이 발송됨.                              ║
-- ║                                                                    ║
-- ║ 안전성: 신규 컬럼 추가 + 기존 데이터 백필 (손실 0).                ║
-- ╚════════════════════════════════════════════════════════════════════╝

-- 1) 컬럼 추가 (3개 테이블 동일 패턴)
ALTER TABLE offline_sales  ADD COLUMN IF NOT EXISTS review_promised_type text;
ALTER TABLE repairs        ADD COLUMN IF NOT EXISTS review_promised_type text;
ALTER TABLE consultations  ADD COLUMN IF NOT EXISTS review_promised_type text;

-- 2) 값 제약 (purchase / repair / consult / NULL)
ALTER TABLE offline_sales
  DROP CONSTRAINT IF EXISTS offline_sales_review_promised_type_chk;
ALTER TABLE offline_sales
  ADD CONSTRAINT offline_sales_review_promised_type_chk
  CHECK (review_promised_type IN ('purchase','repair','consult') OR review_promised_type IS NULL);

ALTER TABLE repairs
  DROP CONSTRAINT IF EXISTS repairs_review_promised_type_chk;
ALTER TABLE repairs
  ADD CONSTRAINT repairs_review_promised_type_chk
  CHECK (review_promised_type IN ('purchase','repair','consult') OR review_promised_type IS NULL);

ALTER TABLE consultations
  DROP CONSTRAINT IF EXISTS consultations_review_promised_type_chk;
ALTER TABLE consultations
  ADD CONSTRAINT consultations_review_promised_type_chk
  CHECK (review_promised_type IN ('purchase','repair','consult') OR review_promised_type IS NULL);

-- 3) 기존 데이터 백필 — source별 자연 매핑
--    (review_promised_at IS NOT NULL 이지만 type 미설정인 행만)
UPDATE offline_sales
   SET review_promised_type = 'purchase'
 WHERE review_promised_at IS NOT NULL
   AND review_promised_type IS NULL;

UPDATE repairs
   SET review_promised_type = 'repair'
 WHERE review_promised_at IS NOT NULL
   AND review_promised_type IS NULL;

UPDATE consultations
   SET review_promised_type = 'consult'
 WHERE review_promised_at IS NOT NULL
   AND review_promised_type IS NULL;

-- 4) 검증 쿼리 (실행 후 결과 확인용 — 사장님이 SQL Editor에서 함께 실행)
-- SELECT COUNT(*) FILTER (WHERE review_promised_at IS NOT NULL) AS promised,
--        COUNT(*) FILTER (WHERE review_promised_type IS NOT NULL) AS typed
-- FROM offline_sales;
-- → 두 숫자가 같아야 정상 백필
