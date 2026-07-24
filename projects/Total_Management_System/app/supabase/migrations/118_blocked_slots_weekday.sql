-- 118: 시간차단에 '매주 반복(요일)' 지원 — 2026-07-24
--
-- 배경: blocked_time_slots 는 date + start_time/end_time 만 있어 "매주 화요일 오후 차단" 같은
--   반복 규칙을 표현할 수 없었다(매번 날짜마다 손으로 등록). weekday 를 추가해 한 줄로 반복.
--
-- 설계:
--   · date IS NOT NULL, weekday IS NULL   → 그 날짜만 차단 (기존 데이터 전부 이 형태)
--   · date IS NULL,     weekday 0..6      → 매주 그 요일 반복 차단 (신규)
--   weekday 는 0=일 … 6=토 (consultation_settings.disabled_weekdays 와 동일 규칙)
--
-- 🔒 하위호환: 기존 소비처는 `.in('date', dates)` / `.eq('date', date)` 로 조회하므로
--    date IS NULL 인 반복행은 자연히 매칭되지 않는다 → 기존 동작 그대로.
--    반복 적용은 소비처(상담/복원수리 슬롯 API)에 요일 조회를 명시적으로 추가해서만 동작한다.
--
-- 되돌리기:
--   ALTER TABLE blocked_time_slots DROP CONSTRAINT blocked_slot_date_or_weekday;
--   ALTER TABLE blocked_time_slots DROP COLUMN weekday;
--   ALTER TABLE blocked_time_slots ALTER COLUMN date SET NOT NULL;

ALTER TABLE blocked_time_slots
  ADD COLUMN IF NOT EXISTS weekday INT;

-- 반복행은 날짜가 없으므로 NOT NULL 해제
ALTER TABLE blocked_time_slots
  ALTER COLUMN date DROP NOT NULL;

-- 날짜형이거나 요일형이거나 — 둘 중 정확히 하나
ALTER TABLE blocked_time_slots
  DROP CONSTRAINT IF EXISTS blocked_slot_date_or_weekday;
ALTER TABLE blocked_time_slots
  ADD CONSTRAINT blocked_slot_date_or_weekday CHECK (
    (date IS NOT NULL AND weekday IS NULL)
    OR (date IS NULL AND weekday BETWEEN 0 AND 6)
  );

CREATE INDEX IF NOT EXISTS idx_blocked_time_slots_weekday ON blocked_time_slots(weekday);

COMMENT ON COLUMN blocked_time_slots.weekday
  IS '매주 반복 차단 요일 (0=일 … 6=토). NULL 이면 date 기반 단발 차단 (118, 2026-07-24)';
