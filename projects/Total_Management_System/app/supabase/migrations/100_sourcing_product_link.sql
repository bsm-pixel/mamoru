-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 100_sourcing_product_link — 소싱 품목 ↔ 제품/부자재 연결            ║
-- ║                                                                    ║
-- ║ 배경 (2026-06-01 사장님): 소싱에서 고른 샘플을 제품/부자재로 바로    ║
-- ║   등록하거나(미등록), 기존 제품에 연결. 제품/부자재 상세에서 소싱     ║
-- ║   정보(업체명·회사링크·품목링크·단가)를 보려고 함.                    ║
-- ║   ※ 부자재 = products(category='SUP') 이므로 연결 FK 하나로 공용.    ║
-- ║                                                                    ║
-- ║ ⚠️ ADDITIVE ONLY — nullable·기본 null. 진행 중 소싱 입력에 영향 0.   ║
-- ╚════════════════════════════════════════════════════════════════════╝

ALTER TABLE sourcing_items
  ADD COLUMN IF NOT EXISTS linked_product_id uuid
  REFERENCES products(id) ON DELETE SET NULL;  -- 제품 삭제돼도 소싱행 보존

CREATE INDEX IF NOT EXISTS idx_sourcing_items_linked_product
  ON sourcing_items(linked_product_id);
