-- 008: 전자 계약서 테이블
-- R6: 서명 캔버스 + 제품 선택 + 결제 정보 + PDF/알림톡

CREATE TABLE IF NOT EXISTS contracts (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  contract_number text NOT NULL UNIQUE,            -- CT-20260226-001 형식
  customer_id uuid REFERENCES customers(id),
  customer_name text NOT NULL,
  customer_phone text,
  customer_email text,
  customer_address text,

  -- 결제 정보
  total_amount integer NOT NULL DEFAULT 0,
  discount_amount integer NOT NULL DEFAULT 0,
  final_amount integer NOT NULL DEFAULT 0,
  payment_method text NOT NULL DEFAULT 'card',     -- card | cash | transfer | mixed
  installment_months integer DEFAULT 0,            -- 할부 개월수 (0 = 일시불)

  -- 서명
  signature_data text,                             -- base64 서명 이미지
  signed_at timestamptz,

  -- PDF/이미지
  pdf_url text,
  image_url text,

  -- 알림톡
  notification_sent_at timestamptz,

  -- 상태
  status text NOT NULL DEFAULT 'draft',            -- draft | signed | sent | completed | cancelled
  memo text,

  -- 이카운트 연동
  ecount_sync_status text NOT NULL DEFAULT 'pending',
  ecount_slip_no text,

  -- 연결: 오프라인 판매로 전환 시
  offline_sale_id uuid REFERENCES offline_sales(id),

  created_by uuid REFERENCES profiles(id),
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

-- 계약서 항목
CREATE TABLE IF NOT EXISTS contract_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  contract_id uuid NOT NULL REFERENCES contracts(id) ON DELETE CASCADE,
  product_id uuid REFERENCES products(id),
  product_name text NOT NULL,
  sku text,
  quantity integer NOT NULL DEFAULT 1,
  unit_price integer NOT NULL DEFAULT 0,
  total_price integer NOT NULL DEFAULT 0,
  option_text text                                 -- 옵션 (사이즈, 각인 등)
);

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_contracts_date ON contracts(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_contracts_customer ON contracts(customer_id);
CREATE INDEX IF NOT EXISTS idx_contracts_status ON contracts(status);
CREATE INDEX IF NOT EXISTS idx_contract_items_contract ON contract_items(contract_id);

-- updated_at 트리거
CREATE OR REPLACE FUNCTION update_contracts_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_contracts_updated_at ON contracts;
CREATE TRIGGER trg_contracts_updated_at
  BEFORE UPDATE ON contracts
  FOR EACH ROW
  EXECUTE FUNCTION update_contracts_updated_at();

-- RLS
ALTER TABLE contracts ENABLE ROW LEVEL SECURITY;
ALTER TABLE contract_items ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can manage contracts"
  ON contracts FOR ALL TO authenticated USING (true) WITH CHECK (true);

CREATE POLICY "Authenticated users can manage contract_items"
  ON contract_items FOR ALL TO authenticated USING (true) WITH CHECK (true);
