-- ============================================================================
-- 086: product_serial_audit_log — 시리얼 이동 이력 (Phase C)
-- ============================================================================
-- 시리얼은 사장님이 가장 민감한 데이터. 누가/언제/무엇을 변경했는지 추적 ledger.
-- DB 트리거 기반(append-only) → 변경 호출처(API 6곳 + RPC) 모두 자동 캡처.
-- 시리얼 무결성 절대 원칙(memory/feedback_serial_integrity_strict.md) 안전망 역할.
-- ============================================================================

-- 1) 이력 테이블 — FK 없음 (시리얼 삭제돼도 이력 보존, append-only 정신)
CREATE TABLE IF NOT EXISTS product_serial_audit_log (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),

  -- 시리얼 식별 (FK 없음 — 시리얼 삭제 후에도 이력 추적 가능)
  serial_id UUID NOT NULL,
  serial_number TEXT,            -- 삭제된 시리얼도 번호로 추적 가능하도록 스냅숏

  -- 액션 종류
  action TEXT NOT NULL CHECK (action IN ('INSERT', 'UPDATE', 'DELETE')),

  -- 추적 대상 6개 필드 (이전/이후 값)
  old_status            TEXT,
  new_status            TEXT,
  old_warehouse_zone    TEXT,
  new_warehouse_zone    TEXT,
  old_sale_item_id      UUID,
  new_sale_item_id      UUID,
  old_offline_sale_id   UUID,
  new_offline_sale_id   UUID,
  old_contract_id       UUID,
  new_contract_id       UUID,
  old_product_id        UUID,
  new_product_id        UUID,

  -- 메타
  changed_by UUID REFERENCES auth.users(id) ON DELETE SET NULL,
  changed_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

-- 2) 조회 최적화 인덱스
CREATE INDEX IF NOT EXISTS idx_psal_serial_id      ON product_serial_audit_log(serial_id, changed_at DESC);
CREATE INDEX IF NOT EXISTS idx_psal_serial_number  ON product_serial_audit_log(serial_number);
CREATE INDEX IF NOT EXISTS idx_psal_changed_at     ON product_serial_audit_log(changed_at DESC);
CREATE INDEX IF NOT EXISTS idx_psal_changed_by     ON product_serial_audit_log(changed_by, changed_at DESC);

-- 3) 트리거 함수 — append-only INSERT 만 수행 (product_serials 무결성 영향 0)
CREATE OR REPLACE FUNCTION log_product_serial_change() RETURNS TRIGGER
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
  IF TG_OP = 'INSERT' THEN
    INSERT INTO product_serial_audit_log (
      serial_id, serial_number, action,
      new_status, new_warehouse_zone, new_sale_item_id,
      new_offline_sale_id, new_contract_id, new_product_id,
      changed_by
    ) VALUES (
      NEW.id, NEW.serial_number, 'INSERT',
      NEW.status, NEW.warehouse_zone, NEW.sale_item_id,
      NEW.offline_sale_id, NEW.contract_id, NEW.product_id,
      auth.uid()
    );
    RETURN NEW;

  ELSIF TG_OP = 'UPDATE' THEN
    -- 추적 6개 필드 중 하나라도 변경됐을 때만 기록 (updated_at만 바뀐 경우 skip)
    IF (OLD.status, OLD.warehouse_zone, OLD.sale_item_id,
        OLD.offline_sale_id, OLD.contract_id, OLD.product_id) IS DISTINCT FROM
       (NEW.status, NEW.warehouse_zone, NEW.sale_item_id,
        NEW.offline_sale_id, NEW.contract_id, NEW.product_id) THEN
      INSERT INTO product_serial_audit_log (
        serial_id, serial_number, action,
        old_status, new_status,
        old_warehouse_zone, new_warehouse_zone,
        old_sale_item_id, new_sale_item_id,
        old_offline_sale_id, new_offline_sale_id,
        old_contract_id, new_contract_id,
        old_product_id, new_product_id,
        changed_by
      ) VALUES (
        NEW.id, NEW.serial_number, 'UPDATE',
        OLD.status, NEW.status,
        OLD.warehouse_zone, NEW.warehouse_zone,
        OLD.sale_item_id, NEW.sale_item_id,
        OLD.offline_sale_id, NEW.offline_sale_id,
        OLD.contract_id, NEW.contract_id,
        OLD.product_id, NEW.product_id,
        auth.uid()
      );
    END IF;
    RETURN NEW;

  ELSIF TG_OP = 'DELETE' THEN
    INSERT INTO product_serial_audit_log (
      serial_id, serial_number, action,
      old_status, old_warehouse_zone, old_sale_item_id,
      old_offline_sale_id, old_contract_id, old_product_id,
      changed_by
    ) VALUES (
      OLD.id, OLD.serial_number, 'DELETE',
      OLD.status, OLD.warehouse_zone, OLD.sale_item_id,
      OLD.offline_sale_id, OLD.contract_id, OLD.product_id,
      auth.uid()
    );
    RETURN OLD;
  END IF;

  RETURN NULL;
END;
$$;

-- 4) 트리거 — AFTER로 product_serials 트랜잭션 영향 없게
DROP TRIGGER IF EXISTS trg_product_serial_audit_log ON product_serials;
CREATE TRIGGER trg_product_serial_audit_log
AFTER INSERT OR UPDATE OR DELETE ON product_serials
FOR EACH ROW EXECUTE FUNCTION log_product_serial_change();

-- 5) RLS — append-only 강제 (SELECT만 허용, INSERT는 트리거가 SECURITY DEFINER 로 우회)
ALTER TABLE product_serial_audit_log ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS psal_select_authenticated ON product_serial_audit_log;
CREATE POLICY psal_select_authenticated ON product_serial_audit_log
  FOR SELECT TO authenticated
  USING (true);

-- INSERT/UPDATE/DELETE 정책 없음 = 모두 거부 (트리거 SECURITY DEFINER만 INSERT 가능)
-- 이로써 애플리케이션이나 사장님도 직접 ledger 조작 불가 → 위변조 방지

-- 6) 검증 쿼리 (코멘트) — 실행 후 확인용
-- SELECT count(*) FROM product_serial_audit_log;  -- 처음에는 0 (이후 시리얼 변경 시 자동 누적)
-- SELECT * FROM product_serial_audit_log ORDER BY changed_at DESC LIMIT 10;
