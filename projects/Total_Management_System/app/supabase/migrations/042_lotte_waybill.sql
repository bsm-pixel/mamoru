-- 042_lotte_waybill.sql — 롯데택배 송장번호 관리 (GAS PROP 대체)

CREATE TABLE IF NOT EXISTS lotte_waybill_config (
  id TEXT PRIMARY KEY DEFAULT 'default',
  start_number BIGINT NOT NULL,
  end_number BIGINT NOT NULL,
  current_number BIGINT NOT NULL,
  updated_at TIMESTAMPTZ DEFAULT now()
);

COMMENT ON TABLE lotte_waybill_config IS '롯데택배 ALPS 송장번호 범위 관리';
COMMENT ON COLUMN lotte_waybill_config.start_number IS '송장번호 시작 (11자리 base)';
COMMENT ON COLUMN lotte_waybill_config.end_number IS '송장번호 종료 (11자리 base)';
COMMENT ON COLUMN lotte_waybill_config.current_number IS '다음 사용할 번호 (11자리 base)';

-- 공휴일 테이블 (복원수리 수거일 선택용)
CREATE TABLE IF NOT EXISTS holidays (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  date DATE NOT NULL UNIQUE,
  name TEXT NOT NULL,
  year INT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_holidays_year ON holidays(year);

COMMENT ON TABLE holidays IS '대한민국 공휴일 (복원수리 수거일 선택에서 제외용)';
