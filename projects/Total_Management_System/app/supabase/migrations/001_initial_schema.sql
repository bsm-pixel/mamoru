-- ============================================
-- MAMORU TMS Phase 1 — 초기 스키마
-- ============================================

-- ENUM 타입
CREATE TYPE order_status AS ENUM (
  'pay_wait', 'pay_done', 'preparing', 'shipping',
  'delivered', 'confirmed', 'cancelled', 'refund_request', 'refunded'
);

CREATE TYPE user_role AS ENUM ('owner', 'staff', 'director');
CREATE TYPE customer_source AS ENUM ('imweb', 'consultation', 'as', 'manual');

-- ============================================
-- profiles: 사용자 프로필
-- ============================================
CREATE TABLE profiles (
  id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  display_name TEXT,
  role user_role NOT NULL DEFAULT 'staff',
  phone TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE profiles ENABLE ROW LEVEL SECURITY;

CREATE POLICY "profiles_read_own" ON profiles
  FOR SELECT USING (auth.uid() = id);

CREATE POLICY "profiles_update_own" ON profiles
  FOR UPDATE USING (auth.uid() = id);

-- 자동 프로필 생성 트리거
CREATE OR REPLACE FUNCTION handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO profiles (id, display_name, role)
  VALUES (NEW.id, COALESCE(NEW.raw_user_meta_data->>'display_name', NEW.email), 'staff');
  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE FUNCTION handle_new_user();

-- ============================================
-- customers: 통합 고객 DB
-- ============================================
CREATE TABLE customers (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  phone TEXT,
  phone_normalized TEXT GENERATED ALWAYS AS (regexp_replace(phone, '\D', '', 'g')) STORED,
  email TEXT,
  postcode TEXT,
  address_road TEXT,
  address_detail TEXT,
  source customer_source NOT NULL DEFAULT 'imweb',
  total_orders INT NOT NULL DEFAULT 0,
  total_spent BIGINT NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_customers_phone ON customers(phone_normalized);
CREATE INDEX idx_customers_name ON customers(name);

ALTER TABLE customers ENABLE ROW LEVEL SECURITY;

CREATE POLICY "customers_all_authenticated" ON customers
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- products: 제품 카탈로그
-- ============================================
CREATE TABLE products (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  sku TEXT UNIQUE NOT NULL,
  name TEXT NOT NULL,
  category TEXT NOT NULL,  -- BL/TH/LO/SL
  price BIGINT NOT NULL DEFAULT 0,
  image_url TEXT,
  tags JSONB,
  stock_quantity INT NOT NULL DEFAULT 0,
  barcode TEXT,
  is_active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE products ENABLE ROW LEVEL SECURITY;

CREATE POLICY "products_read_authenticated" ON products
  FOR SELECT USING (auth.role() = 'authenticated');

CREATE POLICY "products_write_owner" ON products
  FOR ALL USING (
    EXISTS (SELECT 1 FROM profiles WHERE id = auth.uid() AND role = 'owner')
  );

-- ============================================
-- orders: 아임웹 주문
-- ============================================
CREATE TABLE orders (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  imweb_order_no TEXT UNIQUE NOT NULL,
  imweb_order_id TEXT,
  customer_id UUID REFERENCES customers(id) ON DELETE SET NULL,
  orderer_name TEXT NOT NULL DEFAULT '',
  orderer_phone TEXT,
  orderer_email TEXT,
  recipient_name TEXT NOT NULL DEFAULT '',
  recipient_phone TEXT,
  recipient_postcode TEXT,
  recipient_address TEXT,
  recipient_address_detail TEXT,
  recipient_memo TEXT,
  total_price BIGINT NOT NULL DEFAULT 0,
  delivery_fee BIGINT NOT NULL DEFAULT 0,
  discount_amount BIGINT NOT NULL DEFAULT 0,
  paid_amount BIGINT NOT NULL DEFAULT 0,
  payment_method TEXT,
  paid_at TIMESTAMPTZ,
  courier_code TEXT,
  courier_name TEXT,
  invoice_number TEXT,
  shipped_at TIMESTAMPTZ,
  delivered_at TIMESTAMPTZ,
  status order_status NOT NULL DEFAULT 'pay_wait',
  imweb_raw JSONB,
  synced_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  ordered_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_orders_status ON orders(status);
CREATE INDEX idx_orders_customer ON orders(customer_id);
CREATE INDEX idx_orders_ordered ON orders(ordered_at DESC);
CREATE INDEX idx_orders_imweb_no ON orders(imweb_order_no);

ALTER TABLE orders ENABLE ROW LEVEL SECURITY;

CREATE POLICY "orders_all_authenticated" ON orders
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- order_items: 주문 품목
-- ============================================
CREATE TABLE order_items (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  order_id UUID NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
  product_id UUID REFERENCES products(id) ON DELETE SET NULL,
  imweb_product_no TEXT,
  product_name TEXT NOT NULL DEFAULT '',
  option_text TEXT,
  quantity INT NOT NULL DEFAULT 1,
  unit_price BIGINT NOT NULL DEFAULT 0,
  total_price BIGINT NOT NULL DEFAULT 0
);

CREATE INDEX idx_order_items_order ON order_items(order_id);

ALTER TABLE order_items ENABLE ROW LEVEL SECURITY;

CREATE POLICY "order_items_all_authenticated" ON order_items
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- sync_log: 동기화 이력
-- ============================================
CREATE TABLE sync_log (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  sync_type TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'running',
  records_synced INT NOT NULL DEFAULT 0,
  error_message TEXT,
  started_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  completed_at TIMESTAMPTZ
);

ALTER TABLE sync_log ENABLE ROW LEVEL SECURITY;

CREATE POLICY "sync_log_all_authenticated" ON sync_log
  FOR ALL USING (auth.role() = 'authenticated');

-- ============================================
-- waybill_counter: 송장번호 카운터 (싱글턴)
-- ============================================
CREATE TABLE waybill_counter (
  id INT PRIMARY KEY DEFAULT 1 CHECK (id = 1),
  range_start BIGINT NOT NULL,
  range_end BIGINT NOT NULL,
  current_value BIGINT NOT NULL
);

ALTER TABLE waybill_counter ENABLE ROW LEVEL SECURITY;

CREATE POLICY "waybill_counter_all_authenticated" ON waybill_counter
  FOR ALL USING (auth.role() = 'authenticated');

-- 원자적 송장번호 발급 함수
CREATE OR REPLACE FUNCTION next_waybill()
RETURNS TEXT AS $$
DECLARE
  base11 BIGINT;
  check_digit INT;
  inv TEXT;
BEGIN
  UPDATE waybill_counter
  SET current_value = current_value + 1
  WHERE id = 1 AND current_value <= range_end
  RETURNING current_value - 1 INTO base11;

  IF base11 IS NULL THEN
    RAISE EXCEPTION '운송장 대역 소진: 추가 대역 요청 필요';
  END IF;

  check_digit := base11 % 7;
  inv := base11::TEXT || check_digit::TEXT;
  RETURN inv;
END;
$$ LANGUAGE plpgsql;

-- ============================================
-- updated_at 자동 갱신 트리거
-- ============================================
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_profiles_updated BEFORE UPDATE ON profiles
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_customers_updated BEFORE UPDATE ON customers
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_products_updated BEFORE UPDATE ON products
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
CREATE TRIGGER trg_orders_updated BEFORE UPDATE ON orders
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();
