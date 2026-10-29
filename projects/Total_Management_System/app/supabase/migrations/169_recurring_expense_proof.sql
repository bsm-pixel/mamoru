-- 169: 고정 경비(매달 반복)에 증빙 (2026-10-29)
--
-- 배경(사장님): 임대료처럼 매달 같은 돈이 1~3월 등 통째로 빠져 간편장부 「남은 돈」이 부풀었다.
--   고정 경비를 「시작 월」과 함께 등록하면 빠진 달을 한 번에 채우고, 이후엔 매일 크론이 이번 달 분을 자동 등록한다.
--   (lib/expenses/recurring.ts) 만들어지는 경비에 증빙이 「미정」으로 쌓이지 않도록 항목에 증빙을 둔다.

ALTER TABLE recurring_expenses ADD COLUMN IF NOT EXISTS proof_type text
  CHECK (proof_type IN ('card', 'tax_invoice', 'cash_receipt', 'receipt', 'none'));

-- 확인용
SELECT column_name, data_type FROM information_schema.columns
 WHERE table_name = 'recurring_expenses' ORDER BY ordinal_position;
