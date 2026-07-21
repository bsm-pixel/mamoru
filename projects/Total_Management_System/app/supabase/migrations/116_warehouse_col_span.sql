-- 116: 가로로 긴 칸막이(칸 병합) 지원 — col_span (2026-07-20)
--
-- 배경: 렉 한 단에서 실제로는 칸막이를 가로로 길게 써서 C1·D1 두 칸 폭을
--   한 칸으로 쓰는 경우가 있다. 지금은 모든 칸이 1열 폭이라 이 구조를 표현 못 함.
--
-- 설계: col_span(가로로 몇 열을 차지하는가) 추가.
--   · col_span = 1  →  보통 칸 (기존 전부)
--   · col_span = 2  →  bin_no 열부터 오른쪽으로 2열 폭 (C1 이 C·D 를 덮음)
--   병합하면 오른쪽 칸(D1)은 삭제되고 그 자리 제품은 왼쪽(C1)으로 옮겨온다.
--   분리하면 col_span=1 로 되돌리고, 비워진 열은 편집화면에서 점선 빈자리로 다시 생성 가능.
--
-- ⚠️ 재고 수량 로직 무관 — 제품→자리 참조(location_id)만 옮길 뿐 수량 컬럼은 안 건드린다.
--
-- 되돌리기: ALTER TABLE warehouse_locations DROP COLUMN col_span;

ALTER TABLE warehouse_locations
  ADD COLUMN IF NOT EXISTS col_span int NOT NULL DEFAULT 1;

COMMENT ON COLUMN warehouse_locations.col_span
  IS '가로로 차지하는 열 수(1-based). 1=보통 칸, 2 이상=가로 병합된 넓은 칸 (116, 2026-07-20)';

-- 안전값 보정 (혹시 0/NULL 이 섞이면 1로)
UPDATE warehouse_locations SET col_span = 1 WHERE col_span IS NULL OR col_span < 1;
