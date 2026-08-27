-- 135_backfill_exchange_out_os20260820001.sql — 소급 교환건 교환출고 버튼 활성화 (2026-08-27)
-- OS-20260820-001 은 '교환 출고 송장' 기능(마이그134) 이전에 교환돼서 returns.new_product_* 가 비어
-- /returns 상세에 '교환 출고 송장 발행' 버튼이 안 뜬다. 그 판매의 실제 교환품(A2-68FS)·시리얼을 읽어 보강.
-- 안전: 실데이터를 읽어 처리, 못 찾으면 NOTICE만(무해). 재실행해도 UPDATE만 반복(idempotent-ish).

DO $$
DECLARE
  v_sale   offline_sales%ROWTYPE;
  v_item   RECORD;
  v_serial text;
  v_ret_id uuid;
  v_rt     text;
BEGIN
  SELECT * INTO v_sale FROM offline_sales WHERE sale_number = 'OS-20260820-001' LIMIT 1;
  IF v_sale.id IS NULL THEN RAISE NOTICE '[backfill] 판매 OS-20260820-001 못 찾음'; RETURN; END IF;

  -- 교환된 새 제품(A2-68FS) 품목 — sku/이름으로 매칭
  SELECT * INTO v_item FROM offline_sale_items
   WHERE sale_id = v_sale.id AND (sku ILIKE '%A2-68FS%' OR product_name ILIKE '%A2-68FS%')
   LIMIT 1;
  IF v_item.product_id IS NULL THEN
    RAISE NOTICE '[backfill] OS-20260820-001 에서 A2-68FS 품목 못 찾음 (현재 품목 확인 필요)';
    RETURN;
  END IF;

  -- 배정된 새 시리얼 (판매에 sold 로 붙은 A2-68FS 시리얼)
  SELECT serial_number INTO v_serial FROM product_serials
   WHERE offline_sale_id = v_sale.id AND product_id = v_item.product_id AND status = 'sold'
   ORDER BY sold_at DESC NULLS LAST LIMIT 1;

  -- 이미 이 판매의 교환 반품레코드가 있으면 새제품 정보만 보강
  SELECT id INTO v_ret_id FROM returns WHERE sale_id = v_sale.id AND return_type = 'exchange' LIMIT 1;

  IF v_ret_id IS NOT NULL THEN
    UPDATE returns
       SET new_product_id   = v_item.product_id,
           new_product_name = v_item.product_name,
           new_serial_number = COALESCE(new_serial_number, v_serial)
     WHERE id = v_ret_id;
    RAISE NOTICE '[backfill] 기존 반품레코드 보강 완료 id=% (제품=%, 시리얼=%)', v_ret_id, v_item.product_name, v_serial;
  ELSE
    -- 반품레코드가 없으면(매장/구버전 교환) 새로 생성 → /returns 에 뜨고 교환출고 버튼 노출
    v_rt := 'RT-' || to_char(now(), 'YYYYMMDD') || '-' ||
            lpad(((SELECT count(*) FROM returns WHERE return_number LIKE 'RT-' || to_char(now(), 'YYYYMMDD') || '-%') + 1)::text, 3, '0');
    INSERT INTO returns (return_number, return_type, sale_id, product_name,
                         new_product_id, new_product_name, new_serial_number, qty,
                         customer_id, name, phone, pickup_method, status, requested_at, reason, created_by)
    VALUES (v_rt, 'exchange', v_sale.id, 'A2-65FS (교환 전)',
            v_item.product_id, v_item.product_name, v_serial, 1,
            v_sale.customer_id, v_sale.customer_name, v_sale.customer_phone,
            '택배수거', 'requested', now(), '제품 교환(소급 등록)', v_sale.created_by);
    RAISE NOTICE '[backfill] 반품레코드 신규 생성 %  (제품=%, 시리얼=%)', v_rt, v_item.product_name, v_serial;
  END IF;
END $$;
