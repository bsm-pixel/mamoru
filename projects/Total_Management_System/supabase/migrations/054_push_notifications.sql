-- 054: 푸시 알림 테이블 (Supabase Realtime 기반)
-- 서버에서 INSERT → 클라이언트가 Realtime 구독으로 즉시 감지
-- Supabase SQL Editor에서 실행

CREATE TABLE IF NOT EXISTS push_notifications (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  title TEXT NOT NULL,
  body TEXT NOT NULL,
  url TEXT DEFAULT '/dashboard',
  tag TEXT DEFAULT 'mamoru',
  read BOOLEAN DEFAULT false,
  created_at TIMESTAMPTZ DEFAULT now()
);

ALTER TABLE push_notifications ENABLE ROW LEVEL SECURITY;
CREATE POLICY "auth_push_notif" ON push_notifications FOR ALL USING (auth.role() = 'authenticated');

-- Realtime 활성화
ALTER PUBLICATION supabase_realtime ADD TABLE push_notifications;
