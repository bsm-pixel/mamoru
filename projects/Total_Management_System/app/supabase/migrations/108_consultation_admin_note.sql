-- 108: 상담자 메모 (관리자 전용)
-- 고객이 접수 시 넣는 memo와 별개로, 관리자가 직접 기입하는 참고 메모.
-- 고객 비노출(공개 API는 명시 컬럼 select라 자동 제외). TMS 상담 상세 + 구글캘린더 설명란에 반영.
-- repairs.admin_note 와 동일 명명.

ALTER TABLE consultations ADD COLUMN IF NOT EXISTS admin_note TEXT;

COMMENT ON COLUMN consultations.admin_note IS '상담자(관리자) 전용 메모. 고객 비노출. TMS 상세·구글캘린더 설명란 반영. (고객 입력 memo와 별개)';
