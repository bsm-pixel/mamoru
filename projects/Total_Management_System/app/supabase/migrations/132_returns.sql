-- 132_returns.sql — 반품·교환수거(returns) 상태머신 테이블 (2026-08-25, 교환/반품 Phase 2)
-- 배송 교환/반품 시 "구 제품 회수(수거→입고)" 프로세스를 추적. 복원수리(repairs) 상태머신 패턴을 반품용으로 복제.
-- 매장 직접 교환(Phase 1)은 즉시 처리라 이 테이블 미사용. 배송 교환/반품환불만 여기서 추적.

CREATE TABLE IF NOT EXISTS returns (
  id                  uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  return_number       text UNIQUE NOT NULL,               -- 'RT-YYYYMMDD-NNN'
  return_type         text NOT NULL DEFAULT 'exchange',   -- exchange(교환) | refund(반품환불)

  -- 원 판매/제품 연결
  sale_id             uuid REFERENCES offline_sales(id) ON DELETE SET NULL,
  product_id          uuid,
  product_name        text,
  serial_id           uuid,                               -- 반품 대상 시리얼(있으면)
  serial_number       text,
  qty                 integer DEFAULT 1,

  -- 고객
  customer_id         uuid REFERENCES customers(id) ON DELETE SET NULL,
  name                text,
  phone               text,
  phone_normalized    text GENERATED ALWAYS AS (regexp_replace(coalesce(phone, ''), '[^0-9]', '', 'g')) STORED,
  postcode            text,
  address             text,
  address_detail      text,

  -- 수거방식 (repairs.proceed_type 패턴)
  pickup_method       text,                               -- '방문수거' | '택배수거' | '직접반납'
  pickup_date         date,                               -- 캘린더 return_pickup 연동 키
  courier_name        text,                               -- 수거 송장(고객→매장 역방향)
  invoice_number      text,

  -- 상태머신: requested → pickup_scheduled → inbound → inspected → completed / cancelled
  status              text NOT NULL DEFAULT 'requested',
  requested_at        timestamptz DEFAULT now(),
  pickup_scheduled_at timestamptz,
  inbound_at          timestamptz,                        -- 반품 입고완료(구 시리얼 반품창고 확정 시점)
  inspected_at        timestamptz,
  completed_at        timestamptz,
  cancelled_at        timestamptz,
  cancelled_reason    text,

  -- 반품환불 금액
  refund_amount       integer DEFAULT 0,
  refund_method       text,                               -- transfer | card | cash

  -- 사유/메모
  reason              text,
  memo                text,
  admin_note          text,

  -- 알림 중복방지 (make-webhook 연동)
  received_notified_at timestamptz,
  inbound_notified_at  timestamptz,

  -- 감사
  created_by          uuid,
  created_at          timestamptz DEFAULT now(),
  updated_at          timestamptz DEFAULT now()
);

COMMENT ON TABLE returns IS '반품·교환수거 상태머신. 배송 교환/반품환불 시 구 제품 회수(수거→입고→검수→완료) 추적';

-- updated_at 자동 갱신 (001_initial_schema.sql의 전역 함수 재사용)
DROP TRIGGER IF EXISTS trg_returns_updated ON returns;
CREATE TRIGGER trg_returns_updated BEFORE UPDATE ON returns
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_returns_status      ON returns(status);
CREATE INDEX IF NOT EXISTS idx_returns_sale        ON returns(sale_id);
CREATE INDEX IF NOT EXISTS idx_returns_customer    ON returns(customer_id);
CREATE INDEX IF NOT EXISTS idx_returns_pickup_date ON returns(pickup_date);
CREATE INDEX IF NOT EXISTS idx_returns_created      ON returns(created_at DESC);

-- RLS (운영 데이터 — 인증 사용자 전체 접근)
ALTER TABLE returns ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "returns_all" ON returns;
CREATE POLICY "returns_all" ON returns FOR ALL USING (true) WITH CHECK (true);
