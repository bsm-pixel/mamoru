-- 163: 수리내역서 v3 — 그림체 상처 맵 + 확대 사진
-- 변경: 사진 위 손 핀(photo_marks) → 가위 종류별 표준 그림 위 체크 항목(diagram_marks)
--   diagram_marks : 체크 항목 id 배열 (예: ["z-tip-dull","z-mid-dull","z-pivot-tension"])
--                   NULL = v2(사진 핀) 내역서 → 고객 페이지는 옛 화면 그대로. [] = v3 인데 상처 없음(양호)
--                   항목 정의 SSOT = app/src/lib/repair/wound-catalog.ts
--   zoom_photos   : 확대 사진(현미경 등) URL 배열, 가위당 0~3장
-- photo_marks 는 비파괴 보존(구 내역서 표시용)
ALTER TABLE repair_inspections ADD COLUMN IF NOT EXISTS diagram_marks jsonb;
ALTER TABLE repair_inspections ADD COLUMN IF NOT EXISTS zoom_photos jsonb NOT NULL DEFAULT '[]'::jsonb;
