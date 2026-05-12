-- 078: Hub Stats RPC — 매출 3분할 (B2C 제품 / B2B 제품 / 복원수리 전체) 지원
--
-- 배경 (2단계 — 사장님 합의 2026-05-12):
--   메인 대시보드 총매출 = B2C 제품 + B2B 제품 + 복원수리 전체 로 3분할 표시.
--   RS(복원수리) 항목은 제품 매출에서 빠지고 "복원수리 전체"(A 접수 + B 판매RS + C 납품RS)로만 집계.
--   077 까지는 sales.monthAmount 가 "offline_sales(RS 포함) + deliveries(RS 포함)" 한 덩어리였고
--   클라이언트 후처리에서 deliveries 를 더하고 제품매출 B2C/B2B 분리를 직접 쿼리했음.
--   → 이 마이그레이션은 그 분리 계산을 RPC 로 옮겨 후처리 추가 쿼리를 제거하는 최적화.
--
-- 변경 (077 대비):
--   * 'sales' 객체 확장 — 077 의 monthCount/monthAmount(offline_sales 만) → 아래 4개:
--     1. monthCount  = offline_sales 건수 + deliveries 건수 (RS 포함 — "오프라인 판매" 카드 호환용)
--     2. monthAmount = offline_sales(total−discount) + deliveries(total−discount) 전체 (RS 포함 — 호환용)
--     3. salesB2C    = offline_sales(소매/온라인: customer_type IS NULL OR NOT IN dealer/academy) (total−discount) − 그 채널 RS_total
--     4. salesB2B    = offline_sales(딜러/아카데미) (total−discount) − RS_total
--                      + deliveries 전체 (total−discount) − 납품 RS_total      (납품은 전부 B2B 거래처)
--   * 가정: RS 항목엔 할인 미적용 → 제품 매출 = (total_amount − discount_amount) − 그 주문 RS items total_price 합
--   * RS 차감 서브쿼리 필터는 077 의 monthRepair* 패턴과 동일 (osi.category='RS' AND osi.total_price>0 AND os.cancelled_at IS NULL)
--     — offline_sales 합계는 returned_at IS NULL 까지 (077 'sales' 와 동일). 둘의 미세 불일치(반품된 매출의 RS)는 사실상 발생 없음.
--   * deliveries 필터 = 077 monthRepairAmount C채널과 동일: status IN ('confirmed','shipped','settled') AND cancelled_at IS NULL
--   * 그 외 orders / consultations / repairs 객체는 077 과 100% 동일 (1단계 검증 완료 값 보존)
--
-- 클라이언트(use-dashboard-stats.ts useHubStats)는 sales.salesB2C/salesB2B 존재 여부로 077/078 분기 — 미배포여도 정상 동작.

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
        -- ─── 접수시스템(A채널) 매출만 — 월 목표 계산용 (077: 당시 sales.monthAmount 에 B/C 이미 포함이라 A만 추가했음.
        --      078 부터는 대시보드 KPI 가 salesB2C+salesB2B+monthRepairAmount 로 바뀌므로 monthRepairAOnly 는 호환/참고용으로만 유지) ───
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
    -- ─── 매출 객체 — 077 의 monthCount/monthAmount(offline 만) → deliveries 포함 + B2C/B2B 제품매출 분리 ───
    'sales', json_build_object(
      -- 호환: 오프라인판매(RS 포함) + 납품(RS 포함) 전체 (대시보드 "오프라인 판매" 카드용)
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
      -- B2C 제품 매출 = 소매/온라인 offline_sales (total−discount) − 그 채널 RS_total
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
      -- B2B 제품 매출 = 딜러/아카데미 offline_sales (total−discount) − RS_total + 납품 전체 (total−discount) − 납품 RS_total
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
