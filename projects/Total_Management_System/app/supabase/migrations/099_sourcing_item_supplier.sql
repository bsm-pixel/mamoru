-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 099_sourcing_item_supplier — 소싱 품목에 업체(회사) 정보 추가         ║
-- ║                                                                    ║
-- ║ 배경 (2026-05-31 사장님): 가위집 등 샘플을 A·B·C 여러 업체에서 매입.  ║
-- ║   "어느 공장이 좋은 제품을 많이 내놓나"로 업체를 선정하려면 품목마다  ║
-- ║   회사명·회사링크가 있어야 함 (제품 링크만으론 같은 업체인지 헷갈림). ║
-- ║   → 회사명으로 그룹핑해 업체별 채택 현황을 본다.                      ║
-- ╚════════════════════════════════════════════════════════════════════╝

ALTER TABLE sourcing_items ADD COLUMN IF NOT EXISTS supplier_name text; -- 회사명(중문 가능)
ALTER TABLE sourcing_items ADD COLUMN IF NOT EXISTS supplier_url  text; -- 1688 회사 홈

CREATE INDEX IF NOT EXISTS idx_sourcing_items_supplier ON sourcing_items(supplier_name);
