-- 055: 대시보드 할일 메모
-- Supabase SQL Editor에서 실행

CREATE TABLE IF NOT EXISTS dashboard_todos (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  text TEXT NOT NULL,
  completed BOOLEAN DEFAULT false,
  created_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE dashboard_todos ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_todos" ON dashboard_todos FOR ALL USING (auth.role() = 'authenticated');
