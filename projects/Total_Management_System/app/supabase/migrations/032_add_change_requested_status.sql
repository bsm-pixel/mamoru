-- consultation_status enum에 change_requested 추가
-- 고객이 출장 확정 후 일정 변경을 요청할 때 사용
ALTER TYPE consultation_status ADD VALUE IF NOT EXISTS 'change_requested';
