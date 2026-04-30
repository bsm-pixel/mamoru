-- 075: Hub Stats RPC v3
-- 사장님 보고 (2026-04-30 심야 +3) — 두 가지 핵심 버그 fix:
--   1) needAction에 pending_admin 잘못 포함 → 신규 상담이 "일정 재요청"으로 중복 카운트
--   2) monthRepairAmount의 A채널 paid_at 조건 → 사장님 합의 (옵션 A "발생 기준") 적용 → 미입금도 매출로 카운트
--   3) B채널 product_name ILIKE '%복원수리%' → category='RS'로 정확화 + total_price > 0 (무상 0원 제외)
-- 추가: monthRepairMamoru / monthRepairOther / monthRepairB2B 분리 (RPC에 빠져 화면에 0/0으로 표시되던 문제)

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
  -- 075: 접수시스템 고정단가 (마모루 1만 / 타사 2만)
  price_mamoru int := 10000;
  price_other int := 20000;
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
        -- 075: pending_admin 제거 — 신규 상담이 "일정 재요청"으로 중복 카운트되던 버그 fix
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
        -- 075: 옵션 A "발생 기준" — paid_at 조건 제거. A채널 + B채널 합산 (RS catalog)
        'monthRepairAmount', COALESCE((
          -- A채널: 접수시스템 (이번달 접수, 취소 제외) — 고정단가 × 수량
          SELECT SUM(qty_mamoru * price_mamoru + qty_other * price_other) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          -- B채널: 판매시스템 category='RS' (정확화 + 0원 무상 제외)
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
        -- 075 신설: 마모루/타사/B2B 분리 (이전엔 RPC에 없어 화면에 0으로 표시되던 문제 fix)
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
