-- 076: Hub Stats RPC — KST timezone fix
-- 사장님 보고 (2026-05-01 첫날 94% 표시 버그):
--   원인: 075의 now() = timestamptz (UTC). KST 5/1 00:30 → UTC 4/30 15:30 →
--         date_trunc('month', now()) = 4/1 → 4월 한 달치 매출이 5월에 합산
--   해결: now()를 Asia/Seoul로 변환한 후 월/주 시작 계산
--
-- 함수 본체는 075와 100% 동일 — DECLARE/BEGIN의 시점 변수 계산만 KST 캐스팅.
-- 클라이언트 코드 변경 불필요 (동일 함수명, 동일 반환 형식).

CREATE OR REPLACE FUNCTION get_hub_stats()
RETURNS json
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  result json;
  -- 076: KST 기준으로 시점 계산 (UTC 기반 now()의 월/주 경계 오류 fix)
  now_kst timestamp := (now() AT TIME ZONE 'Asia/Seoul');
  mon_iso timestamptz;
  month_iso timestamptz;
  month_start_date date;
  dow int;
  -- 075: 접수시스템 고정단가 (마모루 1만 / 타사 2만)
  price_mamoru int := 10000;
  price_other int := 20000;
BEGIN
  dow := EXTRACT(ISODOW FROM now_kst)::int;
  -- KST 벽시계 기준으로 자른 후 다시 KST 절대시각으로 해석
  mon_iso := (date_trunc('day', now_kst) - ((dow - 1) || ' days')::interval) AT TIME ZONE 'Asia/Seoul';
  month_iso := date_trunc('month', now_kst) AT TIME ZONE 'Asia/Seoul';
  month_start_date := date_trunc('month', now_kst)::date;

  SELECT json_build_object(
    'orders', (
      SELECT json_build_object(
        'payDone', COALESCE(SUM(CASE WHEN status = 'pay_done' THEN 1 ELSE 0 END), 0),
        'preparing', COALESCE(SUM(CASE WHEN status = 'preparing' THEN 1 ELSE 0 END), 0),
        'shipping', COALESCE(SUM(CASE WHEN status = 'shipping' THEN 1 ELSE 0 END), 0),
        'delivered', COALESCE(SUM(CASE WHEN status = 'delivered' THEN 1 ELSE 0 END), 0),
        'weekAmount', COALESCE((
          SELECT SUM(paid_amount) FROM orders
          WHERE ordered_at >= mon_iso AND status NOT IN ('cancelled', 'refunded')
        ), 0),
        'monthAmount', COALESCE((
          SELECT SUM(paid_amount) FROM orders
          WHERE ordered_at >= month_iso AND status NOT IN ('cancelled', 'refunded')
        ), 0)
      )
      FROM orders
      WHERE status IN ('pay_done', 'preparing', 'shipping', 'delivered')
    ),
    'consultations', (
      SELECT json_build_object(
        'newIntake', COALESCE(SUM(CASE WHEN status = 'pending_admin' THEN 1 ELSE 0 END), 0),
        'confirmed', COALESCE(SUM(CASE WHEN status = 'confirmed' THEN 1 ELSE 0 END), 0),
        'needAction', COALESCE(SUM(
          CASE WHEN status IN ('reschedule_requested', 'change_requested')
               AND consultation_type IN ('field_request', 'talk_consult')
          THEN 1 ELSE 0 END
        ), 0)
      )
      FROM consultations
    ),
    'repairs', (
      SELECT json_build_object(
        'intakeNew', COALESCE(SUM(CASE WHEN status = 'intake' AND confirmed_at IS NULL THEN 1 ELSE 0 END), 0),
        'pendingInbound', COALESCE(SUM(
          CASE WHEN status IN ('intake', 'pickup_scheduled')
               AND NOT (status = 'intake' AND confirmed_at IS NULL)
          THEN 1 ELSE 0 END
        ), 0),
        'workingCount', COALESCE(SUM(CASE WHEN status IN ('cost_notified', 'repairing', 'ready_to_ship') THEN 1 ELSE 0 END), 0),
        'workingQty', COALESCE(SUM(
          CASE WHEN status IN ('cost_notified', 'repairing', 'ready_to_ship')
          THEN qty_mamoru + qty_other ELSE 0 END
        ), 0),
        'readyToShip', COALESCE(SUM(CASE WHEN status = 'ready_to_ship' THEN 1 ELSE 0 END), 0),
        'weekRepairTotal', COALESCE((
          SELECT COUNT(*) FROM repairs
          WHERE status IN ('shipped', 'delivered', 'completed') AND shipped_at >= mon_iso
        ), 0),
        'weekRepairMamoru', COALESCE((
          SELECT SUM(qty_mamoru) FROM repairs
          WHERE status IN ('shipped', 'delivered', 'completed') AND shipped_at >= mon_iso
        ), 0),
        'weekRepairOther', COALESCE((
          SELECT SUM(qty_other) FROM repairs
          WHERE status IN ('shipped', 'delivered', 'completed') AND shipped_at >= mon_iso
        ), 0),
        'monthRepairAmount', COALESCE((
          SELECT SUM(qty_mamoru * price_mamoru + qty_other * price_other) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          SELECT SUM(osi.total_price) FROM offline_sale_items osi
          INNER JOIN offline_sales os ON os.id = osi.sale_id
          WHERE osi.category = 'RS'
            AND osi.total_price > 0
            AND os.sale_date >= month_start_date
            AND os.cancelled_at IS NULL
        ), 0),
        'monthRepairCount', COALESCE((
          SELECT COUNT(*) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          SELECT COUNT(*) FROM offline_sale_items osi
          INNER JOIN offline_sales os ON os.id = osi.sale_id
          WHERE osi.category = 'RS'
            AND os.sale_date >= month_start_date
            AND os.cancelled_at IS NULL
        ), 0),
        'monthRepairMamoru', json_build_object(
          'amount', COALESCE((
            SELECT SUM(qty_mamoru * price_mamoru) FROM repairs
            WHERE created_at >= month_iso AND status != 'cancelled'
          ), 0) + COALESCE((
            SELECT SUM(osi.total_price) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND (os.customer_type IS NULL OR os.customer_type NOT IN ('dealer', 'academy'))
              AND osi.product_name NOT ILIKE '%타사%'
          ), 0),
          'count', COALESCE((
            SELECT SUM(qty_mamoru) FROM repairs
            WHERE created_at >= month_iso AND status != 'cancelled'
          ), 0) + COALESCE((
            SELECT SUM(osi.quantity) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND (os.customer_type IS NULL OR os.customer_type NOT IN ('dealer', 'academy'))
              AND osi.product_name NOT ILIKE '%타사%'
          ), 0)
        ),
        'monthRepairOther', json_build_object(
          'amount', COALESCE((
            SELECT SUM(qty_other * price_other) FROM repairs
            WHERE created_at >= month_iso AND status != 'cancelled'
          ), 0) + COALESCE((
            SELECT SUM(osi.total_price) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND (os.customer_type IS NULL OR os.customer_type NOT IN ('dealer', 'academy'))
              AND osi.product_name ILIKE '%타사%'
          ), 0),
          'count', COALESCE((
            SELECT SUM(qty_other) FROM repairs
            WHERE created_at >= month_iso AND status != 'cancelled'
          ), 0) + COALESCE((
            SELECT SUM(osi.quantity) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND (os.customer_type IS NULL OR os.customer_type NOT IN ('dealer', 'academy'))
              AND osi.product_name ILIKE '%타사%'
          ), 0)
        ),
        'monthRepairB2B', json_build_object(
          'amount', COALESCE((
            SELECT SUM(osi.total_price) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND os.customer_type IN ('dealer', 'academy')
          ), 0),
          'count', COALESCE((
            SELECT SUM(osi.quantity) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND os.customer_type IN ('dealer', 'academy')
          ), 0)
        )
      )
      FROM repairs
    ),
    'sales', (
      SELECT json_build_object(
        'monthCount', COALESCE(COUNT(*), 0),
        'monthAmount', COALESCE(SUM(total_amount), 0)
      )
      FROM offline_sales
      WHERE sale_date >= month_start_date
        AND cancelled_at IS NULL
        AND returned_at IS NULL
    )
  ) INTO result;

  RETURN result;
END;
$$;
