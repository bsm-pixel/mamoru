-- 080: get_hub_stats — 복원수리 "자루 수" 카운트에서 배송비 항목 제외 (금액은 그대로 포함)
--
-- 배경 (2026-05-13):
--   복원수리 입력 시 "배송비 3,000원" 항목을 offline_sale_items / delivery_items 에 category='RS', quantity=1 로 저장함.
--   → 금액(매출)은 복원수리 매출에 맞게 들어가는데, "자루 수"(가위 몇 자루 수리) 합산 시 배송비 1줄이 1자루로 잘못 잡혀
--     메인 대시보드 복원수리 카드/회계 복원수리 자루수가 배송비 건당 +1 부풀려졌음.
--   해결: 자루(quantity) 합산 서브쿼리에만 AND product_name <> '배송비' 추가. 금액(amount) 서브쿼리는 그대로(배송비 = 복원수리 매출 포함, 접수시스템 shipping_fee 와 정책 통일).
--
-- 078 대비 변경된 곳: monthRepairCount(offline/delivery qty), monthRepairMamoru.count(offline qty),
--   monthRepairB2B.count(offline + delivery qty) 에 `AND ...product_name <> '배송비'` 추가. 그 외 동일.
--   (monthRepairOther.count 는 product_name ILIKE '%타사%' 라 '배송비' 가 애초에 안 들어가므로 변경 불필요)

CREATE OR REPLACE FUNCTION get_hub_stats()
RETURNS json
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  result json;
  now_kst timestamp := (now() AT TIME ZONE 'Asia/Seoul');
  mon_iso timestamptz;
  month_iso timestamptz;
  month_start_date date;
  dow int;
  price_mamoru int := 10000;
  price_other int := 20000;
BEGIN
  dow := EXTRACT(ISODOW FROM now_kst)::int;
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
        -- ─── 복원수리 매출 전체 = A(repairs) + B(offline RS) + C(delivery RS) — 금액은 배송비 포함 ───
        'monthRepairAmount', COALESCE((
          SELECT SUM(qty_mamoru * price_mamoru + qty_other * price_other) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          SELECT SUM(osi.total_price) FROM offline_sale_items osi
          INNER JOIN offline_sales os ON os.id = osi.sale_id
          WHERE osi.category = 'RS' AND osi.total_price > 0
            AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
        ), 0) + COALESCE((
          SELECT SUM(di.total_price) FROM delivery_items di
          INNER JOIN deliveries dl ON dl.id = di.delivery_id
          WHERE di.category = 'RS' AND di.total_price > 0
            AND dl.delivery_date >= month_start_date
            AND dl.status IN ('confirmed', 'shipped', 'settled')
            AND dl.cancelled_at IS NULL
        ), 0),
        -- ─── 복원수리 수량 전체 (자루) = repairs(qty_m+qty_o) + offline RS(quantity, 배송비 제외) + delivery RS(quantity, 배송비 제외) ───
        'monthRepairCount', COALESCE((
          SELECT SUM(qty_mamoru + qty_other) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          SELECT SUM(osi.quantity) FROM offline_sale_items osi
          INNER JOIN offline_sales os ON os.id = osi.sale_id
          WHERE osi.category = 'RS' AND osi.total_price > 0 AND osi.product_name <> '배송비'
            AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
        ), 0) + COALESCE((
          SELECT SUM(di.quantity) FROM delivery_items di
          INNER JOIN deliveries dl ON dl.id = di.delivery_id
          WHERE di.category = 'RS' AND di.total_price > 0 AND di.product_name <> '배송비'
            AND dl.delivery_date >= month_start_date
            AND dl.status IN ('confirmed', 'shipped', 'settled')
            AND dl.cancelled_at IS NULL
        ), 0),
        -- ─── 접수시스템(A채널) 매출만 — 호환/참고용 ───
        'monthRepairAOnly', COALESCE((
          SELECT SUM(qty_mamoru * price_mamoru + qty_other * price_other) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
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
            WHERE osi.category = 'RS' AND osi.total_price > 0 AND osi.product_name <> '배송비'
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
        -- ─── B2B 복원수리 = 판매시스템 RS(dealer/academy) + 납품 RS ───
        'monthRepairB2B', json_build_object(
          'amount', COALESCE((
            SELECT SUM(osi.total_price) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND os.customer_type IN ('dealer', 'academy')
          ), 0) + COALESCE((
            SELECT SUM(di.total_price) FROM delivery_items di
            INNER JOIN deliveries dl ON dl.id = di.delivery_id
            WHERE di.category = 'RS' AND di.total_price > 0
              AND dl.delivery_date >= month_start_date
              AND dl.status IN ('confirmed', 'shipped', 'settled')
              AND dl.cancelled_at IS NULL
          ), 0),
          'count', COALESCE((
            SELECT SUM(osi.quantity) FROM offline_sale_items osi
            INNER JOIN offline_sales os ON os.id = osi.sale_id
            WHERE osi.category = 'RS' AND osi.total_price > 0 AND osi.product_name <> '배송비'
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND os.customer_type IN ('dealer', 'academy')
          ), 0) + COALESCE((
            SELECT SUM(di.quantity) FROM delivery_items di
            INNER JOIN deliveries dl ON dl.id = di.delivery_id
            WHERE di.category = 'RS' AND di.total_price > 0 AND di.product_name <> '배송비'
              AND dl.delivery_date >= month_start_date
              AND dl.status IN ('confirmed', 'shipped', 'settled')
              AND dl.cancelled_at IS NULL
          ), 0)
        )
      )
      FROM repairs
    ),
    -- ─── 매출 객체 — 078 과 동일 (offline + deliveries 전체 monthCount/monthAmount + 제품매출 B2C/B2B 분리) ───
    'sales', json_build_object(
      'monthCount', (
        COALESCE((SELECT COUNT(*) FROM offline_sales
                  WHERE sale_date >= month_start_date AND cancelled_at IS NULL AND returned_at IS NULL), 0)
        + COALESCE((SELECT COUNT(*) FROM deliveries
                    WHERE delivery_date >= month_start_date AND status IN ('confirmed', 'shipped', 'settled') AND cancelled_at IS NULL), 0)
      ),
      'monthAmount', (
        COALESCE((SELECT SUM(total_amount - COALESCE(discount_amount, 0)) FROM offline_sales
                  WHERE sale_date >= month_start_date AND cancelled_at IS NULL AND returned_at IS NULL), 0)
        + COALESCE((SELECT SUM(total_amount - COALESCE(discount_amount, 0)) FROM deliveries
                    WHERE delivery_date >= month_start_date AND status IN ('confirmed', 'shipped', 'settled') AND cancelled_at IS NULL), 0)
      ),
      'salesB2C', (
        COALESCE((SELECT SUM(os.total_amount - COALESCE(os.discount_amount, 0)) FROM offline_sales os
                  WHERE os.sale_date >= month_start_date AND os.cancelled_at IS NULL AND os.returned_at IS NULL
                    AND (os.customer_type IS NULL OR os.customer_type NOT IN ('dealer', 'academy'))), 0)
        - COALESCE((SELECT SUM(osi.total_price) FROM offline_sale_items osi
                    INNER JOIN offline_sales os ON os.id = osi.sale_id
                    WHERE osi.category = 'RS' AND osi.total_price > 0
                      AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
                      AND (os.customer_type IS NULL OR os.customer_type NOT IN ('dealer', 'academy'))), 0)
      ),
      'salesB2B', (
        COALESCE((SELECT SUM(os.total_amount - COALESCE(os.discount_amount, 0)) FROM offline_sales os
                  WHERE os.sale_date >= month_start_date AND os.cancelled_at IS NULL AND os.returned_at IS NULL
                    AND os.customer_type IN ('dealer', 'academy')), 0)
        - COALESCE((SELECT SUM(osi.total_price) FROM offline_sale_items osi
                    INNER JOIN offline_sales os ON os.id = osi.sale_id
                    WHERE osi.category = 'RS' AND osi.total_price > 0
                      AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
                      AND os.customer_type IN ('dealer', 'academy')), 0)
        + COALESCE((SELECT SUM(dl.total_amount - COALESCE(dl.discount_amount, 0)) FROM deliveries dl
                    WHERE dl.delivery_date >= month_start_date AND dl.status IN ('confirmed', 'shipped', 'settled') AND dl.cancelled_at IS NULL), 0)
        - COALESCE((SELECT SUM(di.total_price) FROM delivery_items di
                    INNER JOIN deliveries dl ON dl.id = di.delivery_id
                    WHERE di.category = 'RS' AND di.total_price > 0
                      AND dl.delivery_date >= month_start_date
                      AND dl.status IN ('confirmed', 'shipped', 'settled') AND dl.cancelled_at IS NULL), 0)
      )
    )
  ) INTO result;

  RETURN result;
END;
$$;
