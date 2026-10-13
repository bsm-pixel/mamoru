-- 161: 은행 입금 알림 기록 — 복원수리 입금확인 자동화 (2026-10-13)
--
-- 배경(사장님): 입금이 들어와도 TMS에서 [입금확인]을 눌러야 고객에게 입금확인 알림톡이 나갔다.
--   우리WON뱅킹 입금 푸시(안드로이드) → MacroDroid → POST /api/bank/deposit 로 받아
--   입금대기 복원수리 건과 「이름+총액」이 정확히 1건 맞으면 자동 입금처리 + 알림톡.
--
-- outcome
--   pending      : 받음, 처리 중
--   matched      : 자동 입금확인 완료 (repair_id)
--   needs_review : 금액 맞는 건은 있으나 확실하지 않음 → 사장님 푸시 (candidate_repair_ids)
--   unmatched    : 복원수리와 무관한 입금 (제품 판매·개인 송금 등) — 기록만
--   already_paid : 매칭됐으나 그 사이 수동 입금처리됨
-- 삭제하지 않는다(이력). 같은 알림 재전송은 dedupe_key 로 막는다.

CREATE TABLE IF NOT EXISTS bank_deposits (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  bank text NOT NULL DEFAULT 'woori',
  depositor text NOT NULL,
  amount integer NOT NULL,
  deposited_at timestamptz NOT NULL,
  raw_text text,
  dedupe_key text NOT NULL UNIQUE,
  outcome text NOT NULL DEFAULT 'pending'
    CHECK (outcome IN ('pending', 'matched', 'needs_review', 'unmatched', 'already_paid')),
  repair_id uuid REFERENCES repairs(id) ON DELETE SET NULL,
  candidate_repair_ids uuid[],
  resolved_at timestamptz,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_bank_deposits_created ON bank_deposits(created_at DESC);

ALTER TABLE bank_deposits ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "bank_deposits_auth_all" ON bank_deposits;
CREATE POLICY "bank_deposits_auth_all" ON bank_deposits
  FOR ALL USING (auth.role() = 'authenticated');

-- 확인용
SELECT column_name, data_type FROM information_schema.columns
 WHERE table_name = 'bank_deposits' ORDER BY ordinal_position;
