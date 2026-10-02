-- 156: 리뷰 이벤트 당첨자 배송 (2026-10-02)
--
-- 흐름: 당첨 발표 → [당첨 안내 알림톡](주소 입력 링크) → 고객이 page_event_address 에서 배송지 확인·저장
--       → TMS [송장 생성](롯데 ALPS) → 기사님 집하 스캔 감지(track-delivery 크론) → [당첨 상품 출고 알림톡]
--
-- 당첨자 1명 = 1행 (reviews.id 기준 유니크). 당첨 마킹 자체는 그대로 reviews.event_month/event_rank 가 SSOT,
-- 이 표는 "배송"만 담당한다. 고객이 저장한 주소는 customers 에도 반영(신규 주소 = 고객 정보 갱신).
-- 삭제 대신 cancelled_at (soft) — 운영 데이터는 상태로 관리.

CREATE TABLE IF NOT EXISTS review_event_shipments (
  id                     uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  review_id              uuid NOT NULL UNIQUE REFERENCES reviews(id) ON DELETE CASCADE,
  event_month            text NOT NULL,              -- 'YYMM' (reviews.event_month 와 동일)
  rank                   int,
  rank_label             text,                       -- '1등' 또는 커스텀 상 이름
  prize                  text,                       -- 당첨 상품명 (발송 시점 스냅샷)
  name                   text NOT NULL,
  phone                  text NOT NULL,              -- 숫자만
  token                  text NOT NULL UNIQUE,       -- 주소 입력 링크용 (추측 불가 32자)
  postcode               text,
  address_road           text,
  address_detail         text,
  delivery_message       text,
  address_prefill_source text,                       -- 링크 첫 진입 때 미리 채운 출처 (customer/repair/order)
  address_submitted_at   timestamptz,                -- 고객이 주소를 확인·저장한 시각
  customer_id            uuid,
  won_notified_at        timestamptz,                -- 당첨 안내 알림톡 발송 시각
  invoice_number         text,
  courier_name           text DEFAULT '롯데택배',
  invoice_created_at     timestamptz,
  shipped_at             timestamptz,                -- 집하(출고) 시각 — 크론 CAS 선점 키
  shipped_notified_at    timestamptz,                -- 출고 알림톡 발송 시각
  delivered_at           timestamptz,
  cancelled_at           timestamptz,
  memo                   text,
  created_at             timestamptz NOT NULL DEFAULT now(),
  updated_at             timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS review_event_shipments_month_idx ON review_event_shipments (event_month);
-- 크론: 송장 있음 + 아직 집하 전
CREATE INDEX IF NOT EXISTS review_event_shipments_pickup_idx
  ON review_event_shipments (invoice_created_at)
  WHERE invoice_number IS NOT NULL AND shipped_at IS NULL AND cancelled_at IS NULL;

ALTER TABLE review_event_shipments ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "review_event_shipments_auth_all" ON review_event_shipments;
CREATE POLICY "review_event_shipments_auth_all" ON review_event_shipments
  FOR ALL TO authenticated USING (true) WITH CHECK (true);
-- 고객 공개 API(page_event_address)는 서버의 service role 로만 접근 — anon 정책 없음

COMMENT ON TABLE review_event_shipments IS '리뷰 이벤트 당첨자 배송 (주소 입력·송장·집하·알림톡). 마이그 156';

-- 확인용
SELECT column_name FROM information_schema.columns WHERE table_name = 'review_event_shipments' ORDER BY ordinal_position;
