-- ============================================================================
-- 087: swap_serials() RPC — Phase B 시리얼 양방향 교환
-- ============================================================================
-- 시나리오: 판매 A에 시리얼 X 등록, 판매 B에 시리얼 Y 등록 →
--           사장님이 둘이 바뀌어야 함을 발견 → 한 번의 트랜잭션으로 X↔Y 동시 스왑
--
-- 안전장치 5겹:
--   1) 동일 시리얼 거부 (실수 방지)
--   2) SELECT FOR UPDATE 로 두 행 lock (race 방지)
--   3) 둘 다 status='sold' 여야 (in_stock 은 Phase A 단방향 이전으로 처리)
--   4) 둘 다 offline_sale_id 가 있어야 + 서로 다른 판매여야
--   5) product_id 반드시 일치 (다른 제품 스왑은 데이터 손상 — 거부)
--
-- 이력: Phase C 트리거가 두 UPDATE 를 product_serial_audit_log 에 자동 캡처
-- ============================================================================

CREATE OR REPLACE FUNCTION swap_serials(p_serial_a uuid, p_serial_b uuid)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_a product_serials%ROWTYPE;
  v_b product_serials%ROWTYPE;
BEGIN
  -- ── 가드 1: 동일 시리얼 ID 거부
  IF p_serial_a = p_serial_b THEN
    RAISE EXCEPTION 'SAME_SERIAL' USING MESSAGE = '같은 시리얼끼리는 교환할 수 없습니다';
  END IF;

  -- ── 가드 2: 두 행 lock (FOR UPDATE) — 동시성 race 차단
  SELECT * INTO v_a FROM product_serials WHERE id = p_serial_a FOR UPDATE;
  SELECT * INTO v_b FROM product_serials WHERE id = p_serial_b FOR UPDATE;

  IF NOT FOUND OR v_a.id IS NULL THEN
    RAISE EXCEPTION 'SERIAL_NOT_FOUND_A' USING MESSAGE = '시리얼 A를 찾을 수 없습니다';
  END IF;
  IF v_b.id IS NULL THEN
    RAISE EXCEPTION 'SERIAL_NOT_FOUND_B' USING MESSAGE = '시리얼 B를 찾을 수 없습니다';
  END IF;

  -- ── 가드 3: 둘 다 sold 상태여야
  IF v_a.status IS DISTINCT FROM 'sold' OR v_b.status IS DISTINCT FROM 'sold' THEN
    RAISE EXCEPTION 'NOT_SOLD' USING MESSAGE = '두 시리얼 모두 판매완료 상태여야 교환할 수 있습니다';
  END IF;

  -- ── 가드 4: 둘 다 판매 연결 + 서로 다른 판매
  IF v_a.offline_sale_id IS NULL OR v_b.offline_sale_id IS NULL THEN
    RAISE EXCEPTION 'NO_SALE' USING MESSAGE = '판매에 연결된 시리얼만 교환할 수 있습니다';
  END IF;
  IF v_a.offline_sale_id = v_b.offline_sale_id THEN
    RAISE EXCEPTION 'SAME_SALE' USING MESSAGE = '같은 판매 안의 시리얼끼리는 교환이 무의미합니다';
  END IF;

  -- ── 가드 5: product_id 반드시 일치 (다른 제품 스왑은 sale_item.product_id 정합성 깨짐)
  IF v_a.product_id IS DISTINCT FROM v_b.product_id THEN
    RAISE EXCEPTION 'PRODUCT_MISMATCH' USING MESSAGE = '같은 제품의 시리얼끼리만 교환할 수 있습니다';
  END IF;

  -- ── 스왑 실행 — 5개 필드 교환 (status / warehouse_zone / sold_at / previous_zone 은 유지)
  UPDATE product_serials SET
    offline_sale_id = v_b.offline_sale_id,
    sale_item_id    = v_b.sale_item_id,
    sold_to_name    = v_b.sold_to_name,
    sold_to_phone   = v_b.sold_to_phone,
    sold_via        = v_b.sold_via
  WHERE id = p_serial_a;

  UPDATE product_serials SET
    offline_sale_id = v_a.offline_sale_id,
    sale_item_id    = v_a.sale_item_id,
    sold_to_name    = v_a.sold_to_name,
    sold_to_phone   = v_a.sold_to_phone,
    sold_via        = v_a.sold_via
  WHERE id = p_serial_b;

  RETURN jsonb_build_object(
    'success', true,
    'serial_a_id', p_serial_a,
    'serial_b_id', p_serial_b,
    'moved_a_to', v_b.offline_sale_id,
    'moved_b_to', v_a.offline_sale_id
  );
END;
$$;

-- 실행 권한 — authenticated 만 호출 가능
REVOKE ALL ON FUNCTION swap_serials(uuid, uuid) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION swap_serials(uuid, uuid) TO authenticated;

-- 검증 쿼리 (코멘트) — 실행 후 함수 확인용
-- SELECT proname, prosrc FROM pg_proc WHERE proname = 'swap_serials';
