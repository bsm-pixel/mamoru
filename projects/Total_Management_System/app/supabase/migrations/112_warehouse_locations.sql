-- 112: 창고 로케이션(정위치) 관리 — 2026-07-18
--
-- 배경: 재고 위치를 머릿속으로만 관리해 "이 제품 어디 있지?"를 매번 찾아야 함.
--       렉·단·칸에 주소를 부여하고(로케이션 코드), 제품마다 고정된 자리를 지정(정위치 관리)해
--       배치도에서 그림으로 찾을 수 있게 한다.
--
-- ⚠️ 설계 원칙 — 재고 수량 로직은 일절 건드리지 않는다:
--   · 수량을 위치별로 쪼개지 않는다(정석 WMS 의 수량×위치 매트릭스 X).
--     보관 재고는 products.raw_stock 단일 정수라, 쪼개는 순간 raw_stock 이 파생값이 되어
--     판매·납품·매입입고·아임웹동기화·시리얼생성·이벤트전환 등 10곳 이상의 갱신 지점이 전부 영향권.
--     불변식 stock_quantity = raw_stock + COUNT(in_stock 시리얼) 은 DB 가 아니라 앱 코드가 맞추고 있어
--     여기를 흔들면 재고 정합성이 깨진다.
--   · 따라서 '제품(SKU) → 자리 1곳' 참조만 추가한다. 수량 컬럼/계산식 무변경.
--
-- 기존 warehouse_zone(raw/ready/display)과의 관계:
--   zone      = 용도 구역 (보관/판매준비/전시)   ← 기존, 유지
--   location  = 물리적 자리 (몇 번 렉 몇 단 몇 칸) ← 신규
--
-- 되돌리기:
--   ALTER TABLE products DROP COLUMN location_id;
--   DROP TABLE warehouse_locations;

CREATE TABLE IF NOT EXISTS warehouse_locations (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code text NOT NULL UNIQUE,                    -- 자리 주소 'R01-2-A' (자릿수 고정 → 정렬 안정)
  label text,                                   -- 사람이 읽는 이름 '1번렉 중단 A칸'
  rack_no int NOT NULL,                         -- 렉 번호
  level_no int NOT NULL,                        -- 단 (1=상단, 위→아래 — 눈으로 보는 순서와 일치)
  bin_no int,                                   -- 칸 (칸을 안 나누면 NULL)
  zone_type text NOT NULL DEFAULT 'storage',    -- storage | ready | display (기존 warehouse_zone 과 개념 정렬)
  sort_order int NOT NULL DEFAULT 0,
  is_active boolean NOT NULL DEFAULT true,
  memo text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

COMMENT ON TABLE  warehouse_locations IS '창고 로케이션(정위치). 렉·단·칸 물리적 자리 주소 (112, 2026-07-18)';
COMMENT ON COLUMN warehouse_locations.code      IS '자리 주소. 예 R01-2-A = 1번렉 2단 A칸';
COMMENT ON COLUMN warehouse_locations.level_no  IS '단. 1=상단(위→아래 번호)';
COMMENT ON COLUMN warehouse_locations.zone_type IS '용도 구역 storage/ready/display — product_serials.warehouse_zone 과 개념 정렬';

CREATE INDEX IF NOT EXISTS idx_wh_loc_grid   ON warehouse_locations (rack_no, level_no, bin_no);
CREATE INDEX IF NOT EXISTS idx_wh_loc_active ON warehouse_locations (is_active) WHERE is_active = true;

-- 제품의 정위치 (1 제품 → 1 자리). 자리를 지워도 제품이 깨지지 않도록 SET NULL.
ALTER TABLE products
  ADD COLUMN IF NOT EXISTS location_id uuid REFERENCES warehouse_locations(id) ON DELETE SET NULL;

COMMENT ON COLUMN products.location_id IS '정위치(보관 자리). NULL=미지정. 재고 수량과 무관한 위치 참조 (112, 2026-07-18)';

CREATE INDEX IF NOT EXISTS idx_products_location ON products(location_id);
