-- 복합 결제 금액 분리 저장 (카드/현금/이체 각각)
ALTER TABLE offline_sales ADD COLUMN IF NOT EXISTS payment_detail jsonb DEFAULT NULL;

-- 기존 데이터 보정: 단독 결제 건에 대해 payment_detail 채우기 (선택적)
-- UPDATE offline_sales SET payment_detail = jsonb_build_object(payment_method, payment_amount)
-- WHERE payment_detail IS NULL AND payment_method IS NOT NULL AND payment_method != 'mixed';
