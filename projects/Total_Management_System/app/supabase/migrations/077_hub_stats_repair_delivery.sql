-- 077: Hub Stats RPC — 복원수리 매출에 납품(deliveries) RS 항목 포함 + 카운트 수량 기준 통일 + monthRepairAOnly 신규
--
-- 배경 (사장님 보고 2026-05-XX):
--   "B2B 거래에서 +B2B수리 22자루 입력했는데 메인 대시보드 복원수리 카드에 1건만 뜨고 B2B 매출 누락"
--   원인: 복원수리 매출 집계가 A(접수시스템 repairs) + B(판매시스템 offline_sale_items RS)만 합산.
--         C채널 — B2B 수리 = delivery_items(category='RS') 가 monthRepairAmount/monthRepairCount에 누락.
--         (075/076은 deliveries RS를 monthRepairB2B.count 에만 부분 반영, 금액·전체카운트 미반영)
--
-- 변경:
--   1. monthRepairAmount = A(repairs) + B(offline RS) + C(delivery RS total_price)   ← C 추가
--   2. monthRepairCount  = repairs 수량(qty_mamoru+qty_other) + offline RS 수량(quantity) + delivery RS 수량(quantity)
--                          ← COUNT(*) 레코드 수 → SUM(quantity) 자루 수 기준으로 통일 (사장님 요청: "수량으로 카운트")
--   3. monthRepairAOnly  = repairs 테이블(접수시스템 A채널) 매출만 (qty_mamoru*1만 + qty_other*2만)   ← 신규 필드
--                          : 월 목표 계산에서 sales.monthAmount(B/C채널 RS 이미 포함) + monthRepairAOnly 만 더해야 중복 없음
--   4. monthRepairB2B.amount/.count 에 delivery RS 추가 (납품은 전부 B2B 거래처)
--   5. sales.monthAmount = SUM(total_amount - COALESCE(discount_amount,0))   ← discount 공제 추가 (클라이언트 fallback과 정책 통일)
--
--   A채널 필터는 075/076 그대로 옵션A(발생 기준 — paid_at 조건 없음, 미입금도 매출 계상) 유지.
--   delivery RS 집계 조건: deliveries.status IN ('confirmed','shipped','settled') AND cancelled_at IS NULL,
--                          delivery_items.category = 'RS' AND total_price > 0  (offline RS 의 '0원 무상 제외' 정책과 통일)
--
-- 클라이언트 (use-dashboard-stats.ts useHubStats RPC 후처리)는 이 RPC 미배포 시에도 동일 결과가 나오도록 보완되어 있음.
--   → 이 마이그레이션은 "RPC 후처리 추가 쿼리를 줄이는 최적화" 성격. 미실행해도 화면은 정상 동작.

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
        -- ─── 복원수리 매출 전체 = A(repairs) + B(offline RS) + C(delivery RS) ───
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
        -- ─── 복원수리 수량 전체 (자루) = repairs(qty_m+qty_o) + offline RS(quantity) + delivery RS(quantity) ───
        'monthRepairCount', COALESCE((
          SELECT SUM(qty_mamoru + qty_other) FROM repairs
          WHERE created_at >= month_iso AND status != 'cancelled'
        ), 0) + COALESCE((
          SELECT SUM(osi.quantity) FROM offline_sale_items osi
          INNER JOIN offline_sales os ON os.id = osi.sale_id
          WHERE osi.category = 'RS' AND osi.total_price > 0
            AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
        ), 0) + COALESCE((
          SELECT SUM(di.quantity) FROM delivery_items di
          INNER JOIN deliveries dl ON dl.id = di.delivery_id
          WHERE di.category = 'RS' AND di.total_price > 0
            AND dl.delivery_date >= month_start_date
            AND dl.status IN ('confirmed', 'shipped', 'settled')
            AND dl.cancelled_at IS NULL
        ), 0),
        -- ─── 접수시스템(A채널) 매출만 — 월 목표 계산용 (B/C채널 RS는 sales.monthAmount에 이미 포함) ───
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
        -- ─── B2B 복원수리 = 판매시스템 RS(dealer/academy) + 납품 RS(전부 B2B 거래처) ───
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
            WHERE osi.category = 'RS' AND osi.total_price > 0
              AND os.sale_date >= month_start_date AND os.cancelled_at IS NULL
              AND os.customer_type IN ('dealer', 'academy')
          ), 0) + COALESCE((
            SELECT SUM(di.quantity) FROM delivery_items di
            INNER JOIN deliveries dl ON dl.id = di.delivery_id
            WHERE di.category = 'RS' AND di.total_price > 0
              AND dl.delivery_date >= month_start_date
              AND dl.status IN ('confirmed', 'shipped', 'settled')
              AND dl.cancelled_at IS NULL
          ), 0)
        )
      )
      FROM repairs
    ),
    'sales', (
      SELECT json_build_object(
        'monthCount', COALESCE(COUNT(*), 0),
        'monthAmount', COALESCE(SUM(total_amount - COALESCE(discount_amount, 0)), 0)  -- 077: discount 공제 (클라이언트 fallback과 통일)
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
