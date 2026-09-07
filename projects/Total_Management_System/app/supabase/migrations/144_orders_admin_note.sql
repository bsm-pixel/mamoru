-- 144: 주문 사장님 메모(관리자 전용) — 고객 비노출 작업 메모
-- 복원수리(admin_note)/상담(admin_note)과 동일한 관리자 전용 자유 메모 필드.
-- 고객이 남긴 recipient_memo(배송 요청)와는 별개.
ALTER TABLE orders ADD COLUMN IF NOT EXISTS admin_note TEXT;
