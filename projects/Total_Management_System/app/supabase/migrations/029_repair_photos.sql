-- 029: 복원수리 사진 저장
CREATE TABLE IF NOT EXISTS repair_photos (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  repair_id uuid NOT NULL REFERENCES repairs(id) ON DELETE CASCADE,
  file_path text NOT NULL,          -- Storage 경로: repairs/{repair_id}/{filename}
  file_name text NOT NULL,          -- 원본 파일명
  file_size integer DEFAULT 0,      -- 바이트
  memo text,                        -- 사진 메모 (선택)
  uploaded_by uuid REFERENCES profiles(id),
  created_at timestamptz DEFAULT now()
);

CREATE INDEX idx_repair_photos_repair ON repair_photos(repair_id);

-- RLS: 인증된 사용자 모두 접근 가능
ALTER TABLE repair_photos ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_all_repair_photos" ON repair_photos FOR ALL TO authenticated USING (true) WITH CHECK (true);
