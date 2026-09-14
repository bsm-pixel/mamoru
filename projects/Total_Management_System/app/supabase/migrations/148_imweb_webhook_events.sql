-- 148: 아임웹 웹훅 수신 기록 (2026-09-14)
-- 목적: 아임웹 → TMS 웹훅이 "실제로 들어오는지 / 어떤 값이 오는지"를 원본 그대로 남긴다.
--   · 취소·반품·교환·거절 등 이벤트의 실제 페이로드 확인 → 알림톡(MMR_) 템플릿·매핑 확정 근거
--   · "알림톡 안 왔어요" 문의 시 1차 진단(웹훅 수신 여부)
-- 쓰기: /api/imweb/webhook (service role). 운영 데이터이므로 삭제하지 않고 누적.
-- 코드는 이 테이블이 없어도 기존 주문 동기화가 그대로 동작하도록 기록 실패를 삼킨다.

CREATE TABLE IF NOT EXISTS imweb_webhook_events (
  id              bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  received_at     timestamptz NOT NULL DEFAULT now(),
  event_type      text,                 -- 페이로드 최상위 eventType (예: ORDER_CANCEL_REQUEST)
  order_no        text,                 -- data.orderNo (문자열화)
  action          text NOT NULL,        -- sync(주문 동기화 실행) | logged(기록만) | no_order_no
  payload         jsonb,                -- 원본 그대로 (가공·필터 없음)
  headers         jsonb,                -- 요청 헤더 (cookie/authorization 제외)
  processed_at    timestamptz,          -- sync 완료 시각
  process_result  jsonb,                -- syncSingleOrder 결과
  process_error   text
);

CREATE INDEX IF NOT EXISTS idx_imweb_webhook_events_received ON imweb_webhook_events(received_at DESC);
CREATE INDEX IF NOT EXISTS idx_imweb_webhook_events_type     ON imweb_webhook_events(event_type);
CREATE INDEX IF NOT EXISTS idx_imweb_webhook_events_order    ON imweb_webhook_events(order_no);

-- RLS: 고객 개인정보(주문자명·연락처)가 원본에 포함될 수 있음 → 로그인 사용자 조회만 허용, 쓰기는 service role
ALTER TABLE imweb_webhook_events ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "imweb_webhook_events_select" ON imweb_webhook_events;
CREATE POLICY "imweb_webhook_events_select" ON imweb_webhook_events FOR SELECT TO authenticated USING (true);
