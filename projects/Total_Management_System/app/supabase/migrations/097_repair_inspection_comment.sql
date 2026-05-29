-- 097: 수리내역서 v2 — 가위별 진단 멘트(comment) 컬럼
-- 기존: 진단 코멘트가 repairs.admin_note (건 단위 1개)에만 저장됨
-- 변경: 가위(repair_inspections)별로 진단 및 내역을 독립 저장 → 고객 수리내역서에 가위별 표시
-- photo_marks(jsonb)는 컬럼 변경 없음 — 점/선/플래그를 같은 배열에 저장(선=x2,y2 / 플래그=flag:true)
ALTER TABLE repair_inspections ADD COLUMN IF NOT EXISTS comment text;
