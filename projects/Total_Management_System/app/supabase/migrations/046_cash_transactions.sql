-- 046: 입출금 관리
CREATE TABLE IF NOT EXISTS cash_transactions (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  transaction_date date NOT NULL DEFAULT CURRENT_DATE,
  type text NOT NULL CHECK (type IN ('income', 'expense')),
  category text NOT NULL DEFAULT '기타',
  amount integer NOT NULL DEFAULT 0,
  memo text,
  balance_after integer,
  source_type text,
  source_id uuid,
  created_by uuid,
  created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_cash_tx_date ON cash_transactions(transaction_date);
CREATE INDEX IF NOT EXISTS idx_cash_tx_type ON cash_transactions(type);
