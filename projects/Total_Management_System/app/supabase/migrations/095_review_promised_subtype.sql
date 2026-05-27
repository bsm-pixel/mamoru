-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 095_review_promised_subtype — 리뷰 약속 세부 유형 컬럼 추가         ║
-- ║                                                                    ║
-- ║ 배경 (2026-05-27 사장님 신고 후속):                                ║
-- ║   094 에서 review_promised_type (purchase/repair/consult) 만 저장. ║
-- ║   기존 ReviewRequestModal 은 subtype 분기 보유:                     ║
-- ║     • 복원수리 → direct_visit (직접방문) / pickup (방문수거)         ║
-- ║     • 상담 → store_visit (직접방문) / field_request (출장) / talk_consult (톡상담)║
-- ║     • 제품구매 → subtype 없음                                       ║
-- ║   자동 발송 시 subtype 미전달 → 알림톡 본문/링크 정확도 ↓            ║
-- ║                                                                    ║
-- ║ 개선:                                                              ║
-- ║   review_promised_subtype 컬럼 추가 → 약속 시 사장님이 칩으로 선택   ║
-- ║   → 자동 발송 시 sendReviewRequestNotification 의 subtype 인자로 전달║
-- ║                                                                    ║
-- ║ 안전성: 신규 컬럼 추가 (기존 데이터 손실 0, NULL 허용).             ║
-- ╚════════════════════════════════════════════════════════════════════╝

-- 1) 컬럼 추가 (3개 테이블)
ALTER TABLE offline_sales  ADD COLUMN IF NOT EXISTS review_promised_subtype text;
ALTER TABLE repairs        ADD COLUMN IF NOT EXISTS review_promised_subtype text;
ALTER TABLE consultations  ADD COLUMN IF NOT EXISTS review_promised_subtype text;

-- 2) 값 제약 — type 별로 허용 subtype 다름 (NULL 허용, purchase 일 때는 NULL 권장)
--    repair: direct_visit | pickup
--    consult: store_visit | field_request | talk_consult
--    purchase: NULL
ALTER TABLE offline_sales
  DROP CONSTRAINT IF EXISTS offline_sales_review_promised_subtype_chk;
ALTER TABLE offline_sales
  ADD CONSTRAINT offline_sales_review_promised_subtype_chk
  CHECK (review_promised_subtype IN ('direct_visit','pickup','store_visit','field_request','talk_consult') OR review_promised_subtype IS NULL);

ALTER TABLE repairs
  DROP CONSTRAINT IF EXISTS repairs_review_promised_subtype_chk;
ALTER TABLE repairs
  ADD CONSTRAINT repairs_review_promised_subtype_chk
  CHECK (review_promised_subtype IN ('direct_visit','pickup','store_visit','field_request','talk_consult') OR review_promised_subtype IS NULL);

ALTER TABLE consultations
  DROP CONSTRAINT IF EXISTS consultations_review_promised_subtype_chk;
ALTER TABLE consultations
  ADD CONSTRAINT consultations_review_promised_subtype_chk
  CHECK (review_promised_subtype IN ('direct_visit','pickup','store_visit','field_request','talk_consult') OR review_promised_subtype IS NULL);

-- 3) 기존 데이터 백필 — source + 알려진 컨텍스트 기반 자연 매핑
--    • repairs.proceed_type → repair subtype 매핑
--    • consultations.consultation_type → consult subtype 매핑
--    • offline_sales: subtype 모르므로 NULL 유지 (수동 발송 모달의 디폴트 매핑에 위임)

-- repairs: proceed_type 이 있으면 매핑
UPDATE repairs
   SET review_promised_subtype = CASE
         WHEN proceed_type = '직접방문' OR proceed_type = 'direct_visit' THEN 'direct_visit'
         WHEN proceed_type = '방문수거' OR proceed_type = 'pickup'       THEN 'pickup'
         ELSE NULL
       END
 WHERE review_promised_at IS NOT NULL
   AND review_promised_type = 'repair'
   AND review_promised_subtype IS NULL;

-- consultations: consultation_type 그대로 매핑 (store_visit / field_request / talk_consult)
UPDATE consultations
   SET review_promised_subtype = consultation_type
 WHERE review_promised_at IS NOT NULL
   AND review_promised_type = 'consult'
   AND review_promised_subtype IS NULL
   AND consultation_type IN ('store_visit', 'field_request', 'talk_consult');
