-- 152: EVENT 캠페인 — 무료(증정·체험단) 이벤트 지원 (2026-09-16)
--
-- 문제(사장님 지적): 범용 EVENT 템플릿이 "유료 주문"을 전제로 굳어 있었다.
--   · EVENT_신청완료 본문에 `결제 금액`·`받으실 곳`·`입금 계좌`가 고정 →
--     증정/체험단인데 "결제 금액 0원"이 떠서 고객이 "왜 결제가?" 하게 된다
--   · #{items} 는 항상 "제품명 N개" → 체험단인데 제품명이 노출
--   · 더 근본: 판매 전환이 confirm_payment(입금확인) 한 곳에만 있어
--     무료 이벤트도 [입금확인]을 눌러야 진행되고, 그 순간
--     "입금이 확인되었습니다 / 0원" 알림톡이 나간다
--
-- 해결: 캠페인이 "이 이벤트는 어떤 성격인가"를 들고 있게 하고,
--       고객에게 보이는 문구는 TMS 가 조립한다(알림톡은 변수 중첩이 안 되므로 서버에서 완성).

ALTER TABLE event_campaigns ADD COLUMN IF NOT EXISTS payment_type text NOT NULL DEFAULT 'paid';
ALTER TABLE event_campaigns ADD COLUMN IF NOT EXISTS items_label text;

COMMENT ON COLUMN event_campaigns.payment_type IS
  'paid=입금 필요(기존 흐름) / free=증정·체험단(입금 없음). free 면 금액·계좌 줄 생략 + 입금확인 알림톡 미발송';
COMMENT ON COLUMN event_campaigns.items_label IS
  '알림톡 #{items} 표기 대체 문구(예: 체험단 신청). 비우면 제품명+수량을 그대로 표기. 내부 품목·재고·판매전환에는 영향 없음';

-- 값 방어: 오타로 다른 값이 들어가면 유료로 동작하게(안전한 쪽)
ALTER TABLE event_campaigns DROP CONSTRAINT IF EXISTS event_campaigns_payment_type_chk;
ALTER TABLE event_campaigns ADD CONSTRAINT event_campaigns_payment_type_chk
  CHECK (payment_type IN ('paid', 'free'));

-- 확인용
SELECT id, name, type, payment_type, items_label, customer_notice FROM event_campaigns ORDER BY created_at DESC;
