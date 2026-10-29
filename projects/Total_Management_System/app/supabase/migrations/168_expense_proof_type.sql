-- 168: 경비 증빙 꼬리표 (2026-10-29)
--
-- 배경(사장님): 간이과세자·직접 신고. 종소세(간편장부) 때 "이 경비는 무엇으로 증빙되나"를 한눈에 보고,
--   홈택스에 자동으로 안 잡히는 것(영수증만·증빙 없음)만 따로 챙기려 한다.
-- 값: card=사업자카드(7620 체크 — 홈택스 사업용카드 등록분은 자동) / tax_invoice=세금계산서(전자는 홈택스 자동)
--     cash_receipt=현금영수증(지출증빙) / receipt=영수증만(해외 결제 메일 등 — 원본 보관) / none=증빙 없음
-- NULL = 아직 안 정함(옛 데이터). 간편장부 「증빙 챙길 것」 시트엔 receipt·none·NULL 이 올라간다.

ALTER TABLE expenses             ADD COLUMN IF NOT EXISTS proof_type text CHECK (proof_type IN ('card', 'tax_invoice', 'cash_receipt', 'receipt', 'none'));
ALTER TABLE expense_inbox        ADD COLUMN IF NOT EXISTS proof_type text CHECK (proof_type IN ('card', 'tax_invoice', 'cash_receipt', 'receipt', 'none'));
ALTER TABLE expense_vendor_rules ADD COLUMN IF NOT EXISTS proof_type text CHECK (proof_type IN ('card', 'tax_invoice', 'cash_receipt', 'receipt', 'none'));

-- 메일 대기함: 메일 종류로 기본값 (카드 승인 메일=사업자카드, 세금계산서=세금계산서, 나머지 영수증·청구서=영수증만)
UPDATE expense_inbox
   SET proof_type = CASE doc_type WHEN 'card' THEN 'card' WHEN 'tax_invoice' THEN 'tax_invoice' ELSE 'receipt' END
 WHERE proof_type IS NULL;

-- 메일에서 확인해 만든 경비: 그 메일의 기본값을 그대로
UPDATE expenses e
   SET proof_type = i.proof_type
  FROM expense_inbox i
 WHERE e.source_ref = 'mail:' || i.gmail_id
   AND e.proof_type IS NULL;

-- 확인용
SELECT 'expenses' AS t, proof_type, count(*) FROM expenses GROUP BY proof_type
UNION ALL
SELECT 'expense_inbox', proof_type, count(*) FROM expense_inbox GROUP BY proof_type
ORDER BY 1, 2;
