-- 149: 알림톡 on/off 토글 설정값 삭제 (2026-09-14 · 항상 발송 원칙)
-- ⚠️ 반드시 TMS 배포(토글 판정 코드 제거) "후"에 실행.
--    배포 전 실행 시 구버전 코드가 review.auto_request_on_completion 을 기본값 false 로 읽어 자동 후기요청이 멈춤.
-- 삭제 대상 = 코드에서 더 이상 읽지 않는 키
--   · 고객 알림톡 토글: 전체 on/off + 유형별 7종
--   · 자동 후기요청 토글
--   · 사장님 푸시 토글 push.* (2026-08-01 게이팅 제거 후 무시되던 값)
-- 유지: notifications.webhook_* (Make 웹훅 URL) · notifications.sound_* (앱 알림음)

DELETE FROM system_settings
 WHERE key IN (
   'notifications.master_enabled',
   'notifications.consultation_received',
   'notifications.repair_received',
   'notifications.repair_cost_notice',
   'notifications.repair_payment_confirmed',
   'notifications.repair_shipped',
   'notifications.review_request',
   'notifications.sales_shipped',
   'notifications.event_received',
   'notifications.event_payment_notice',
   'notifications.event_payment_confirmed',
   'notifications.event_shipped',
   'notifications.stock_received',
   'notifications.stock_payment_notice',
   'notifications.stock_payment_confirmed',
   'notifications.return_received',
   'notifications.return_inbound',
   'review.auto_request_on_completion'
 )
    OR key LIKE 'push.%';

-- 확인(0건이어야 함):
-- SELECT key FROM system_settings WHERE key LIKE 'push.%' OR key IN ('notifications.master_enabled','review.auto_request_on_completion');
