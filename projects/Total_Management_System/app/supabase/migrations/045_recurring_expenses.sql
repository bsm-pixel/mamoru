-- 045: 고정 경비 (매월 반복)
CREATE TABLE IF NOT EXISTS recurring_expenses (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  category text NOT NULL,
  amount integer NOT NULL DEFAULT 0,
  memo text,
  is_active boolean NOT NULL DEFAULT true,
  created_at timestamptz NOT NULL DEFAULT now()
);

COMMENT ON TABLE recurring_expenses IS '고정 경비 (임대료, 인건비 등 매월 자동 등록용)';
