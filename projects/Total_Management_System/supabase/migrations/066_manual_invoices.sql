-- 066: 빠른 송장(별도 송장) — 판매와 무관한 ALPS 송장 발급 테이블
-- 매출 KPI 미반영. offline_sales와 분리.
-- Supabase SQL Editor에서 실행

CREATE TABLE IF NOT EXISTS manual_invoices (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),

  -- ALPS 송장
  invoice_number text NOT NULL UNIQUE,

  -- 고객 (FK + 발급 시점 스냅샷 — 고객 정보 변경 후에도 추적 가능)
  customer_id uuid NOT NULL REFERENCES customers(id),
  customer_name text NOT NULL,
  customer_phone text NOT NULL,
  receiver_postcode text NOT NULL,
  receiver_address_road text NOT NULL,
  receiver_address_detail text,

  -- 송장 내용
  goods_name text NOT NULL CHECK (char_length(goods_name) <= 50),
  delivery_message text,

  -- 발급/취소 메타
  created_by uuid REFERENCES profiles(id),
  created_at timestamptz NOT NULL DEFAULT now(),
  cancelled_at timestamptz,
  cancelled_by uuid REFERENCES profiles(id),
  cancelled_reason text,
  alps_cancel_failed boolean NOT NULL DEFAULT false
);

COMMENT ON TABLE manual_invoices IS '빠른 송장(별도 송장) — 판매와 무관한 ALPS 송장 발급. 매출 KPI 미반영';

CREATE INDEX IF NOT EXISTS idx_manual_invoices_created_at_desc
  ON manual_invoices (created_at DESC);

CREATE INDEX IF NOT EXISTS idx_manual_invoices_customer
  ON manual_invoices (customer_id, created_at DESC);

CREATE INDEX IF NOT EXISTS idx_manual_invoices_active
  ON manual_invoices (created_at DESC)
  WHERE cancelled_at IS NULL;

ALTER TABLE manual_invoices ENABLE ROW LEVEL SECURITY;

CREATE POLICY "manual_invoices_authenticated" ON manual_invoices
  FOR ALL TO authenticated USING (true) WITH CHECK (true);
