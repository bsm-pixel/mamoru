-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 098_sourcing — 1688 샘플 소싱 "선별기"                              ║
-- ║                                                                    ║
-- ║ 배경 (2026-05-31 사장님 확정):                                     ║
-- ║   1688에서 신상 샘플을 소량 들여올 때, 뭐가 뭔지 안 놓치게 추적하고  ║
-- ║   실테스트 후 "팔자/아니다"를 선별하는 도구. 역할은                  ║
-- ║   매입(샘플 발주 기록) → 입고 시 제품 매칭 → 실테스트 → 선별 까지.   ║
-- ║                                                                    ║
-- ║ ⚠️ 의도적으로 제외 (기존 매입관리 purchase_orders 와 다름):         ║
-- ║   - 수량/소계: 이 단계 역할은 "선별"이지 수량관리 아님 → 없음        ║
-- ║   - SKU 자동채번 / 판매가: 제품 등록 시 아임웹/제품에서 하는 일      ║
-- ║   - products 테이블 INSERT: 등록은 사장님이 직접(아임웹 먼저→동기화  ║
-- ║     또는 TMS 부가제품 직접등록). 이 도구는 선별 리스트만 출력.       ║
-- ║                                                                    ║
-- ║ 출력 = 채택(selected)된 item 들의 리스트(품목명+1688링크+특징+사진+  ║
-- ║   단가). 사장님이 이 리스트 보고 직접 등록.                          ║
-- ╚════════════════════════════════════════════════════════════════════╝

-- 소싱 발주(샘플 단위 묶음)
CREATE TABLE IF NOT EXISTS sourcing_pos (
  id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  po_number     text NOT NULL UNIQUE,           -- SRC-YYYYMMDD-NNN
  supplier_name text,                            -- 1688 회사명(중문 가능)
  supplier_url  text,                            -- 1688 회사 홈
  order_date    date NOT NULL DEFAULT CURRENT_DATE,
  exchange_rate numeric NOT NULL DEFAULT 195,    -- CNY → KRW (참고용)
  status        text NOT NULL DEFAULT 'sourcing',-- sourcing | done
  memo          text,
  created_by    uuid REFERENCES profiles(id),
  created_at    timestamptz DEFAULT now(),
  updated_at    timestamptz DEFAULT now()
);

-- 소싱 품목(샘플 1종 = 라벨 1장)
CREATE TABLE IF NOT EXISTS sourcing_items (
  id                uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  po_id             uuid NOT NULL REFERENCES sourcing_pos(id) ON DELETE CASCADE,
  sticker_no        text NOT NULL UNIQUE,        -- SRC-YYYYMMDD-NNN-001 (QR 내용)
  vendor_url        text,                         -- 1688 상품 링크 (등록 시 참조)
  product_name      text NOT NULL DEFAULT '',
  features_memo     text,
  unit_price        numeric NOT NULL DEFAULT 0,   -- CNY 단가 (수량 없음 — 선별용)
  moq               integer,                      -- 선택
  inbound_photos    jsonb NOT NULL DEFAULT '[]'::jsonb, -- Storage URL 배열
  inbound_memo      text,
  inspection_status text NOT NULL DEFAULT 'pending', -- pending|matched|selected|rejected
  selected_at       timestamptz,                  -- 채택 시각
  sort_order        integer NOT NULL DEFAULT 0,
  created_at        timestamptz DEFAULT now(),
  updated_at        timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_sourcing_items_po ON sourcing_items(po_id);
CREATE INDEX IF NOT EXISTS idx_sourcing_items_status ON sourcing_items(inspection_status);
CREATE INDEX IF NOT EXISTS idx_sourcing_pos_status ON sourcing_pos(status);

-- RLS: 운영 데이터 — 인증 사용자 전체 접근 (기존 운영 테이블과 동일 정책)
ALTER TABLE sourcing_pos ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "sourcing_pos_all" ON sourcing_pos;
CREATE POLICY "sourcing_pos_all" ON sourcing_pos FOR ALL USING (true) WITH CHECK (true);

ALTER TABLE sourcing_items ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "sourcing_items_all" ON sourcing_items;
CREATE POLICY "sourcing_items_all" ON sourcing_items FOR ALL USING (true) WITH CHECK (true);
