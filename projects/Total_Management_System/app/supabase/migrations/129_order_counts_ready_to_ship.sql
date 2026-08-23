-- 129: get_order_counts 에 배송대기(ready_to_ship) + 입금대기(pay_wait) 카운트 추가
-- ⚠️ 반드시 128(enum ADD VALUE)을 먼저 실행·커밋한 뒤 실행할 것.
--    (함수 본문의 'ready_to_ship' 리터럴이 enum에 존재해야 CREATE 성공)

CREATE OR REPLACE FUNCTION get_order_counts()
RETURNS json LANGUAGE sql STABLE AS $$
  SELECT json_build_object(
    'pay_wait', (SELECT count(*) FROM orders WHERE status = 'pay_wait'),
    'pay_done', (SELECT count(*) FROM orders WHERE status = 'pay_done'),
    'preparing', (SELECT count(*) FROM orders WHERE status = 'preparing'),
    'ready_to_ship', (SELECT count(*) FROM orders WHERE status = 'ready_to_ship'),
    'shipping', (SELECT count(*) FROM orders WHERE status = 'shipping'),
    'delivered', (SELECT count(*) FROM orders WHERE status = 'delivered'),
    'cancel_pending', (SELECT count(*) FROM orders WHERE status = 'cancel_pending'),
    'cancelled', (SELECT count(*) FROM orders WHERE status = 'cancelled')
  );
$$;
