-- 056: Hub Stats RPC v3 — workingCount에서 ready_to_ship 제거 (중복)
-- Supabase SQL Editor에서 실행

CREATE OR REPLACE FUNCTION get_hub_stats()
RETURNS json
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  result json;
  mon_iso timestamptz;
  month_iso timestamptz;
  month_start_date date;
  now_ts timestamptz := now();
  dow int;
BEGIN
  dow := EXTRACT(ISODOW FROM now_ts)::int;
  mon_iso := date_trunc('day', now_ts) - ((dow - 1) || ' days')::interval;
  month_iso := date_trunc('month', now_ts);
  month_start_date := date_trunc('month', now_ts)::date;

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
          CASE WHEN status IN ('reschedule_requested', 'change_requested', 'pending_admin')
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
        -- v3: workingCount에서 ready_to_ship 제거 (readyToShip에서 별도 카운트)
        'workingCount', COALESCE(SUM(CASE WHEN status IN ('cost_notified', 'repairing') THEN 1 ELSE 0 END), 0),
        'workingQty', COALESCE(SUM(
          CASE WHEN status IN ('cost_notified', 'repairing')
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
          SELECT SUM(total_amount) FROM repairs
          WHERE paid_at IS NOT NULL AND created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          SELECT SUM(osi.total_price) FROM offline_sale_items osi
          INNER JOIN offline_sales os ON os.id = osi.sale_id
          WHERE osi.product_name ILIKE '%복원수리%'
            AND os.sale_date >= month_start_date
            AND os.cancelled_at IS NULL
        ), 0),
        'monthRepairCount', COALESCE((
          SELECT COUNT(*) FROM repairs
          WHERE paid_at IS NOT NULL AND created_at >= month_iso AND status != 'cancelled'
        ), 0)
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
