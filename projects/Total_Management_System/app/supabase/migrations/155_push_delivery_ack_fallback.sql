-- 155: 관리자 푸시 "수신 확인 + 문자 예비발송" (2026-09-29)
--
-- 문제(사장님 신고: 로그인이 풀려 있던 동안 접수 알림을 놓쳤다):
--   웹푸시는 브라우저/OS가 토큰을 조용히 끊을 수 있는데, 서버는 "보냈다"까지만 알고
--   **폰에 도착했는지는 몰랐다.** 발송 결과도 콘솔에만 찍혀 몇 건을 놓쳤는지조차 셀 수 없었다.
--
-- 해결:
--   1) push_notifications 에 발송 결과(대상/성공/실패) 기록
--   2) 서비스워커가 알림을 받는 순간 /api/push/ack 로 "받았음" 기록 (휴대폰 수신만 acked_at 인정)
--   3) 2분 안에 휴대폰 수신이 없으면 크론이 사장님 번호로 문자 1통 (fallback_sent_at 으로 1회만)
--
-- 기존 행은 fallback_needed=false(기본값) → 배포 직후 옛 알림에 문자가 몰려 나가지 않는다.

ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS target_count int;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS mobile_target_count int;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS sent_count int;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS failed_count int;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS acked_at timestamptz;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS acked_device text;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS fallback_needed boolean NOT NULL DEFAULT false;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS fallback_sent_at timestamptz;
ALTER TABLE push_notifications ADD COLUMN IF NOT EXISTS fallback_result text;

COMMENT ON COLUMN push_notifications.acked_at IS '휴대폰이 이 알림을 실제로 받은 시각(서비스워커 ack). NULL=미수신';
COMMENT ON COLUMN push_notifications.fallback_needed IS 'true=휴대폰 미수신 시 문자 예비발송 대상';
COMMENT ON COLUMN push_notifications.fallback_sent_at IS '문자 예비발송 시각(선점 표시 겸용 — 1건당 1회만)';

-- 크론이 1분마다 "미수신·미발송" 행만 훑는다
CREATE INDEX IF NOT EXISTS push_notifications_fallback_pending_idx
  ON push_notifications (created_at)
  WHERE fallback_needed = true AND acked_at IS NULL AND fallback_sent_at IS NULL;

-- 기기별 "마지막으로 알림을 실제로 받은 시각" — 설정 화면 표시·끊김 판단용
ALTER TABLE push_subscriptions ADD COLUMN IF NOT EXISTS last_ack_at timestamptz;
COMMENT ON COLUMN push_subscriptions.last_ack_at IS '이 기기가 마지막으로 푸시를 실제 수신한 시각(서비스워커 ack)';

-- 확인용
SELECT column_name FROM information_schema.columns
WHERE table_name IN ('push_notifications', 'push_subscriptions')
  AND column_name IN ('acked_at', 'fallback_needed', 'fallback_sent_at', 'last_ack_at')
ORDER BY column_name;
