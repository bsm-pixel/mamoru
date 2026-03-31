-- 041_consultation_settings.sql — 상담 설정 + 휴무일 (GAS Script Properties 대체)

-- 상담 운영 설정 테이블
CREATE TABLE IF NOT EXISTS consultation_settings (
  id TEXT PRIMARY KEY DEFAULT 'default',
  start_hour INT NOT NULL DEFAULT 10,
  end_hour INT NOT NULL DEFAULT 20,
  duration_min INT NOT NULL DEFAULT 60,
  step_min INT NOT NULL DEFAULT 10,
  disabled_weekdays INT[] NOT NULL DEFAULT '{0}',
  field_buffer_before INT NOT NULL DEFAULT 90,
  field_buffer_after INT NOT NULL DEFAULT 90,
  updated_at TIMESTAMPTZ DEFAULT now()
);

INSERT INTO consultation_settings (id) VALUES ('default')
ON CONFLICT (id) DO NOTHING;

COMMENT ON TABLE consultation_settings IS '상담 운영 설정 (영업시간, 슬롯 간격, 휴무 요일, 출장 버퍼)';
COMMENT ON COLUMN consultation_settings.disabled_weekdays IS '휴무 요일 (0=일, 1=월, ..., 6=토)';
COMMENT ON COLUMN consultation_settings.field_buffer_before IS '출장 예약 전 버퍼 (분)';
COMMENT ON COLUMN consultation_settings.field_buffer_after IS '출장 예약 후 버퍼 (분)';

-- 휴무일 테이블
CREATE TABLE IF NOT EXISTS closed_dates (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  date DATE NOT NULL UNIQUE,
  reason TEXT,
  created_at TIMESTAMPTZ DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_closed_dates_date ON closed_dates(date);

COMMENT ON TABLE closed_dates IS '특정 휴무일 (명절, 임시 휴무 등)';

-- 상담 리마인더 발송 추적 컬럼
ALTER TABLE consultations ADD COLUMN IF NOT EXISTS remind_24h_at TIMESTAMPTZ;
ALTER TABLE consultations ADD COLUMN IF NOT EXISTS remind_2h_at TIMESTAMPTZ;
