-- ╔════════════════════════════════════════════════════════════════════╗
-- ║ 096_blocked_time_slots — 날짜별 30분 단위 시간대 차단                ║
-- ║                                                                    ║
-- ║ 배경 (2026-05-27 사장님 요구):                                     ║
-- ║   달력관리 화면에서 개인 일정/스케줄에 따라 "이 시간부터 이 시간까지 ║
-- ║   예약불가"(30분 단위)를 날짜별로 막고 싶음. 기존 closed_dates 는    ║
-- ║   date UNIQUE 라 전일 휴무만 가능 → 시간대 차단용 별도 테이블 신설.  ║
-- ║                                                                    ║
-- ║ 동작:                                                              ║
-- ║   고객 셀프예약(매장/출장/톡 + 복원수리 직접방문)의 가용 슬롯 계산   ║
-- ║   시 이 시간대를 차단. 사장님 측 흐름(admin-create/suggest/         ║
-- ║   manual-confirm)은 검증 추가 X (블랙아웃 룰).                       ║
-- ║                                                                    ║
-- ║   한 날짜에 여러 시간대 등록 가능 (date UNIQUE 아님).               ║
-- ╚════════════════════════════════════════════════════════════════════╝

CREATE TABLE IF NOT EXISTS blocked_time_slots (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  date DATE NOT NULL,
  start_time TEXT NOT NULL,   -- 'HH:MM' (30분 단위)
  end_time TEXT NOT NULL,     -- 'HH:MM' (30분 단위, start_time 보다 늦음)
  reason TEXT,
  created_at TIMESTAMPTZ DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_blocked_time_slots_date ON blocked_time_slots(date);

-- RLS: 운영 데이터 — 인증 사용자 전체 접근 (기존 closed_dates 와 동일 정책)
ALTER TABLE blocked_time_slots ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "blocked_time_slots_all" ON blocked_time_slots;
CREATE POLICY "blocked_time_slots_all" ON blocked_time_slots
  FOR ALL USING (true) WITH CHECK (true);
