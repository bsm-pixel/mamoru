-- 113: 렉 정보 테이블 — 단마다 칸 수가 다른 실제 렉 구조 지원 (2026-07-18)
--
-- 배경(112 후속): 112 는 "렉의 모든 단이 같은 칸 수"를 가정해 생성했는데,
--   실제 렉은 1단 2칸 / 2단 6칸 / 4단은 칸 없이 선반 통째 처럼 단마다 다르다.
--   DB·화면은 이미 단별로 칸을 따로 저장/렌더하므로 구조 변경은 불필요했고,
--   부족한 건 '렉 전체 열 수'를 담을 자리 하나였다.
--
-- 설계: 렉을 N열 그리드로 본다.
--   · 각 단은 그 N열 중 앞에서부터 몇 칸을 쓰는지가 다름 (나머지는 빈 공간)
--   · bin_no IS NULL 인 로케이션 = 그 단을 칸 없이 '선반 통째'로 쓰는 것 (전 열을 가로지름)
--
-- 되돌리기: DROP TABLE warehouse_racks;

CREATE TABLE IF NOT EXISTS warehouse_racks (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  rack_no int NOT NULL UNIQUE,
  label text,                                  -- '전면 렉' 등 부르는 이름
  columns int NOT NULL DEFAULT 4,              -- 렉 전체 열 수 (그리드 기준선)
  sort_order int NOT NULL DEFAULT 0,
  memo text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

COMMENT ON TABLE  warehouse_racks IS '렉 정보. 렉을 N열 그리드로 보고 단마다 쓰는 칸 수를 다르게 (113, 2026-07-18)';
COMMENT ON COLUMN warehouse_racks.columns IS '렉 전체 열 수. 각 단은 이 중 앞에서부터 일부만 쓸 수 있음';

CREATE INDEX IF NOT EXISTS idx_wh_racks_sort ON warehouse_racks (sort_order, rack_no);

-- 112 로 이미 만든 렉이 있으면 렉 정보를 채워준다 (열 수 = 그 렉에서 가장 칸이 많은 단 기준)
INSERT INTO warehouse_racks (rack_no, columns, sort_order)
SELECT rack_no,
       GREATEST(COALESCE(MAX(bin_no), 1), 1) AS columns,
       rack_no AS sort_order
FROM warehouse_locations
GROUP BY rack_no
ON CONFLICT (rack_no) DO NOTHING;
