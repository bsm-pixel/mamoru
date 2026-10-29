-- 166: 영수증·세금계산서 메일 → 경비 자동 정리 (2026-10-29)
--
-- 배경(사장님): 경비를 손으로 넣다 보니 6월 이후 입력이 멈췄다. 카드 결제 메일·세금계산서·구독 영수증이
--   bsm@mamoru.kr 로 모이고 있으니(다른 계정은 자동 전달), TMS가 매일 메일을 읽어 「확인 대기」에 올리고
--   사장님은 [확인] 한 번으로 경비 확정. 고정 지출(구독·임대) / 사용량 지출(택배·부자재)을 구분한다.
--
-- expense_inbox.status
--   pending   : 메일에서 찾음, 사장님 확인 전 (경비·리포트에 안 잡힘)
--   confirmed : [확인] → expenses 에 1건 생성 (expense_id)
--   ignored   : [무시] (영수증 아님·중복 등)
-- 삭제하지 않는다(이력). 같은 메일은 gmail_id UNIQUE 로 한 번만.
--
-- expense_vendor_rules: 거래처별 기억 — [확인] 때 고른 분류·고정/사용량을 다음 메일에 자동 적용,
--   auto_ignore=true 면 그 거래처 메일은 다음부터 목록에 안 올린다.

-- 1) expenses: 고정/사용량 · 거래처 · 출처
ALTER TABLE expenses ADD COLUMN IF NOT EXISTS kind text CHECK (kind IN ('fixed', 'variable'));
ALTER TABLE expenses ADD COLUMN IF NOT EXISTS vendor text;
ALTER TABLE expenses ADD COLUMN IF NOT EXISTS source_ref text;  -- 'mail:<gmail id>'
-- 기존 고정 경비 일괄 등록분([고정] 메모)은 고정으로
UPDATE expenses SET kind = 'fixed' WHERE kind IS NULL AND memo LIKE '%[고정]%';

-- 2) 메일에서 찾은 영수증 (확인 대기함)
CREATE TABLE IF NOT EXISTS expense_inbox (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  gmail_id text NOT NULL UNIQUE,
  received_at timestamptz,
  sender text,
  subject text,
  vendor text,
  doc_type text NOT NULL DEFAULT 'receipt'
    CHECK (doc_type IN ('card', 'receipt', 'invoice', 'tax_invoice')),
  amount integer,                 -- 원화. 메일에 금액이 없으면(PDF만) NULL → 확인 때 입력
  currency text NOT NULL DEFAULT 'KRW',
  amount_original numeric,        -- 외화 원금 (예: USD 100.00) — 원화 청구액은 확인 때 입력
  expense_date date,
  category text,
  kind text CHECK (kind IN ('fixed', 'variable')),
  memo text,
  snippet text,
  status text NOT NULL DEFAULT 'pending'
    CHECK (status IN ('pending', 'confirmed', 'ignored')),
  expense_id uuid REFERENCES expenses(id) ON DELETE SET NULL,
  decided_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_expense_inbox_status ON expense_inbox(status, received_at DESC);

ALTER TABLE expense_inbox ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "expense_inbox_auth_all" ON expense_inbox;
CREATE POLICY "expense_inbox_auth_all" ON expense_inbox
  FOR ALL USING (auth.role() = 'authenticated');

-- 3) 거래처 기억
CREATE TABLE IF NOT EXISTS expense_vendor_rules (
  vendor text PRIMARY KEY,
  category text,
  kind text CHECK (kind IN ('fixed', 'variable')),
  auto_ignore boolean NOT NULL DEFAULT false,
  updated_at timestamptz NOT NULL DEFAULT now()
);

ALTER TABLE expense_vendor_rules ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "expense_vendor_rules_auth_all" ON expense_vendor_rules;
CREATE POLICY "expense_vendor_rules_auth_all" ON expense_vendor_rules
  FOR ALL USING (auth.role() = 'authenticated');

-- 확인용
SELECT table_name, column_name, data_type FROM information_schema.columns
 WHERE table_name IN ('expense_inbox', 'expense_vendor_rules')
    OR (table_name = 'expenses' AND column_name IN ('kind', 'vendor', 'source_ref'))
 ORDER BY table_name, ordinal_position;
