-- 151: 푸시 구독을 "사용자당 1개" → "기기당 1개" 로 (2026-09-15)
--
-- 문제(사장님 신고: PC·모바일 둘 다 알림이 안 온다):
--   api/push/subscribe 가 single-token-per-user 정책이라
--     DELETE FROM push_subscriptions WHERE user_id = ? AND token <> ?
--   를 실행했다. 즉 **새 기기에서 TMS를 열 때마다 다른 기기의 토큰이 삭제**됐다.
--     · PC에서 열면 → 모바일 토큰 삭제 → 모바일 알림 끊김
--     · 모바일에서 열면 → PC 토큰 삭제 → PC 알림 끊김
--   실측: 2026-09-15 기준 등록 기기 1대뿐.
--
-- 원래 의도는 "같은 기기가 캐시 비움·재구독으로 토큰을 중복 누적하는 것" 차단이었다.
-- 그건 user_id 가 아니라 **기기 단위**로 막아야 맞다 → device_id 도입.
--
--   device_id = 브라우저 localStorage 에 저장되는 UUID (기기+브라우저 단위로 고정)
--   같은 device_id 로 다시 구독하면 그 기기의 옛 토큰만 교체되고, 다른 기기는 살아남는다.

ALTER TABLE push_subscriptions ADD COLUMN IF NOT EXISTS device_id text;
ALTER TABLE push_subscriptions ADD COLUMN IF NOT EXISTS last_seen_at timestamptz;

COMMENT ON COLUMN push_subscriptions.device_id IS
  '기기 식별자(localStorage UUID). 같은 기기의 토큰 교체용. NULL=구버전 등록분';
COMMENT ON COLUMN push_subscriptions.last_seen_at IS
  '이 기기가 마지막으로 TMS를 연 시각. 오래된 기기 정리 판단용';

-- 같은 사용자+같은 기기는 1행만 (device_id 가 있는 행에만 적용 → 구버전 NULL 행은 영향 없음)
CREATE UNIQUE INDEX IF NOT EXISTS push_subscriptions_user_device_uniq
  ON push_subscriptions (user_id, device_id)
  WHERE device_id IS NOT NULL;

-- 확인용
SELECT id, user_id, device_id, device_info, updated_at FROM push_subscriptions ORDER BY updated_at DESC;
