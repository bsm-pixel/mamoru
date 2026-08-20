-- 127: review_promised_subtype 허용값에 'delivery'(택배 수리) 추가
--
-- 배경: 095 제약이 허용 subtype에 'delivery'를 빠뜨림.
--   그런데 앱은 이후(119) repair subtype에 'delivery'(택배 수리)를 추가 → RS(복원수리) 품목이
--   포함된 판매에서 리뷰약속 토글이 type=repair, subtype=delivery 를 저장하려다 CHECK 위반.
--   API가 update 에러를 삼켜 "성공 토스트는 뜨는데 토글은 OFF 고정" 버그로 나타남.
--   → 3개 테이블 제약을 delivery 포함으로 재생성. (앱 코드/API 허용값과 일치)
ALTER TABLE offline_sales DROP CONSTRAINT IF EXISTS offline_sales_review_promised_subtype_chk;
ALTER TABLE offline_sales ADD CONSTRAINT offline_sales_review_promised_subtype_chk
  CHECK (review_promised_subtype IN ('direct_visit','pickup','delivery','store_visit','field_request','talk_consult') OR review_promised_subtype IS NULL);

ALTER TABLE repairs DROP CONSTRAINT IF EXISTS repairs_review_promised_subtype_chk;
ALTER TABLE repairs ADD CONSTRAINT repairs_review_promised_subtype_chk
  CHECK (review_promised_subtype IN ('direct_visit','pickup','delivery','store_visit','field_request','talk_consult') OR review_promised_subtype IS NULL);

ALTER TABLE consultations DROP CONSTRAINT IF EXISTS consultations_review_promised_subtype_chk;
ALTER TABLE consultations ADD CONSTRAINT consultations_review_promised_subtype_chk
  CHECK (review_promised_subtype IN ('direct_visit','pickup','delivery','store_visit','field_request','talk_consult') OR review_promised_subtype IS NULL);
