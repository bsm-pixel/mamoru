-- 170: 우체국 방문접수(수거) 자동 접수 — 예약번호 기록 + 진행 중 잠금 (2026-11-10)
--   TMS 복원수리 상세 [우체국 수거접수] 버튼 → /api/repair/epost-pickup (서울 리전 서버 크롬)
--   epost_reservation_no: 우체국 예약접수번호(16자리). 값이 있으면 버튼 숨김 = 중복 접수 방지
--                         '확인필요' = 신청은 눌렀으나 번호를 못 읽음 → 우체국 예약조회에서 직접 확인
--   epost_requested_at  : 접수 진행 중 표시. 3분 안 재요청(두 번 클릭·다른 기기)은 서버가 거절
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS epost_reservation_no text;
ALTER TABLE repairs ADD COLUMN IF NOT EXISTS epost_requested_at timestamptz;

COMMENT ON COLUMN repairs.epost_reservation_no IS '우체국 방문접수 예약접수번호 (170). 있으면 자동접수 버튼 숨김';
COMMENT ON COLUMN repairs.epost_requested_at IS '우체국 자동접수 진행 중 잠금 시각 (170). 3분 내 재요청 거절';
