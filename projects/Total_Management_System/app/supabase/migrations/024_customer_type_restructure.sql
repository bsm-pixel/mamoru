-- Phase 0: 기존 'b2b' → 'dealer'로 변환 (실제 비즈니스 모델 반영)
-- customer_type 허용값: retail, online, dealer, academy, supplier
UPDATE customers SET customer_type = 'dealer' WHERE customer_type = 'b2b';
