-- 114: 수납함(서랍) 지원 — 한 단 안에 행×열 격자 (2026-07-18)
--
-- 배경(113 후속): 렉 한 단에 '가위 보관 수납함'을 두면 그 안이 6행 10열처럼 격자가 된다.
--   113 까지는 한 단 = 한 줄(열만) 이라 이 구조를 표현할 수 없었다.
--
-- 설계: bin_row(행)를 추가해 한 단을 M행 × N열로 표현한다.
--   bin_no  = 열 (A,B,C… 로 표기)
--   bin_row = 행 (1,2,3… 로 표기)
--
--   단의 세 가지 형태:
--     · 선반 통째    : bin_no NULL, bin_row NULL      → R01-4
--     · 단순 칸막이  : bin_row = 1 (한 줄)             → R01-2-A   (행 번호 안 붙임 = 기존 코드 유지)
--     · 수납함       : bin_row 1..M, bin_no 1..N       → R01-2-A1 … R01-2-J6
--
-- ⚠️ 하위호환: 기존에 만든 칸은 전부 '1행'으로 백필한다. 코드(R01-2-A)는 그대로 유지되므로
--    이미 붙여둔 라벨을 다시 뽑을 필요가 없다.
--
-- 되돌리기: ALTER TABLE warehouse_locations DROP COLUMN bin_row;

ALTER TABLE warehouse_locations
  ADD COLUMN IF NOT EXISTS bin_row int;

COMMENT ON COLUMN warehouse_locations.bin_no  IS '열 번호(1-based). NULL = 칸 없이 선반 통째';
COMMENT ON COLUMN warehouse_locations.bin_row IS '행 번호(1-based). 수납함이면 2 이상. NULL = 선반 (114, 2026-07-18)';

-- 기존 칸은 모두 1행짜리 단순 칸막이였다 → bin_row = 1 백필
UPDATE warehouse_locations
   SET bin_row = 1
 WHERE bin_no IS NOT NULL
   AND bin_row IS NULL;

-- 정렬 키를 행까지 포함하도록 재계산 (렉→단→행→열)
UPDATE warehouse_locations
   SET sort_order = rack_no * 1000000
                  + level_no * 10000
                  + COALESCE(bin_row, 0) * 100
                  + COALESCE(bin_no, 0);

CREATE INDEX IF NOT EXISTS idx_wh_loc_cell
  ON warehouse_locations (rack_no, level_no, bin_row, bin_no);
