-- 027: 성능 최적화 RPC 함수들

-- 1) 주문 상태별 카운트 (6개 쿼리 → 1개)
CREATE OR REPLACE FUNCTION get_order_counts()
RETURNS json LANGUAGE sql STABLE AS $$
  SELECT json_build_object(
    'pay_done', (SELECT count(*) FROM orders WHERE status = 'pay_done'),
    'preparing', (SELECT count(*) FROM orders WHERE status = 'preparing'),
    'shipping', (SELECT count(*) FROM orders WHERE status = 'shipping'),
    'delivered', (SELECT count(*) FROM orders WHERE status = 'delivered'),
    'cancel_pending', (SELECT count(*) FROM orders WHERE status = 'cancel_pending'),
    'cancelled', (SELECT count(*) FROM orders WHERE status = 'cancelled')
  );
$$;

-- 2) 복원수리 탭별 카운트 (8개 쿼리 → 1개)
CREATE OR REPLACE FUNCTION get_repair_tab_counts()
RETURNS json LANGUAGE sql STABLE AS $$
  SELECT json_build_object(
    'intake', (SELECT count(*) FROM repairs WHERE status = 'intake' AND confirmed_at IS NULL),
    'pickup_needed', (SELECT count(*) FROM repairs WHERE status = 'intake' AND proceed_type = '방문수거' AND confirmed_at IS NULL),
    'inbound_waiting', (SELECT count(*) FROM repairs WHERE (status = 'intake' AND confirmed_at IS NOT NULL AND proceed_type != '방문수거') OR status = 'pickup_scheduled'),
    'in_progress', (SELECT count(*) FROM repairs WHERE status IN ('cost_notified', 'repairing')),
    'ready_to_ship', (SELECT count(*) FROM repairs WHERE status = 'ready_to_ship'),
    'shipped', (SELECT count(*) FROM repairs WHERE status IN ('shipped', 'delivered', 'completed'))
  );
$$;
