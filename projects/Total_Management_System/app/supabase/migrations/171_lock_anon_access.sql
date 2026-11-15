-- 171 로그인 없이(공개 키만으로) 읽히던 표 잠그기 (2026-11-15)
--
-- 배경: 브라우저에 내려가는 공개 키(anon)만으로 표 61개 중 19개가 전부 읽혔다
--   (납품·반품·후기 전화번호·현금출납·경비·발주·소싱·단가표·롯데 송장 설정 등).
--   returns·sourcing_*·blocked_time_slots 는 규칙이 "모두 허용"이라 쓰기까지 열려 있었다.
--
-- 조사 결과(코드 전수): 이 표들을 로그인 없이 쓰는 경로는 전부 서버가 서비스 키로 읽는다
--   (예약 달력·후기 공개 목록·크론·웹훅). 고객 페이지(page.mamoru.kr)는 수파베이스를 직접 부르지 않는다.
--   → 공개 키 접근을 막아도 고객 화면·알림톡·자동 작업은 그대로 동작한다.
--
-- 바뀌는 것: 아래 표는 "로그인한 사용자만 읽기·쓰기" 규칙 하나만 남긴다. 자료는 바뀌지 않는다.
-- 되돌리기: 맨 아래 주석 참고.

BEGIN;

-- 1) 표: 행 보안 켜기 + 기존 규칙 전부 지우고 "로그인 사용자 전체 허용" 하나만
DO $$
DECLARE
  t text;
  p record;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'deliveries','delivery_items','returns','reviews',
    'cash_transactions','expenses','recurring_expenses',
    'purchase_orders','purchase_order_items',
    'sourcing_items','sourcing_pos',
    'customer_product_catalog','supplier_product_catalog',
    'lotte_waybill_config',
    'blocked_time_slots','closed_dates','holidays','consultation_settings'
  ]
  LOOP
    -- 표가 아니면(뷰 등) 건너뜀
    IF NOT EXISTS (
      SELECT 1 FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
      WHERE n.nspname = 'public' AND c.relname = t AND c.relkind = 'r'
    ) THEN
      RAISE NOTICE '건너뜀(표 아님 또는 없음): %', t;
      CONTINUE;
    END IF;

    EXECUTE format('ALTER TABLE public.%I ENABLE ROW LEVEL SECURITY', t);

    FOR p IN SELECT policyname FROM pg_policies WHERE schemaname = 'public' AND tablename = t LOOP
      EXECUTE format('DROP POLICY %I ON public.%I', p.policyname, t);
    END LOOP;

    EXECUTE format(
      'CREATE POLICY %I ON public.%I FOR ALL TO authenticated USING (true) WITH CHECK (true)',
      t || '_login_all', t
    );
  END LOOP;
END $$;

-- 2) 뷰: 입고 예정 재고 — 코드에서 쓰지 않음. 공개 키 접근만 뺀다
REVOKE ALL ON public.v_pending_stock FROM anon;

-- 3) 함수: 공개 키로는 못 부르게 (고객 병합·시리얼 맞바꾸기·송장 번호 뽑기·집계 4종)
--    TMS 화면(로그인 사용자)과 서버(서비스 키)는 그대로 부를 수 있다
DO $$
DECLARE
  f record;
BEGIN
  FOR f IN
    SELECT p.oid::regprocedure AS sig
    FROM pg_proc p JOIN pg_namespace n ON n.oid = p.pronamespace
    WHERE n.nspname = 'public'
      AND p.proname IN ('merge_customers','swap_serials','next_waybill',
                        'get_activity_types_by_phones','get_hub_stats','get_order_counts','get_repair_tab_counts')
  LOOP
    EXECUTE format('REVOKE ALL ON FUNCTION %s FROM PUBLIC, anon', f.sig);
    EXECUTE format('GRANT EXECUTE ON FUNCTION %s TO authenticated, service_role', f.sig);
  END LOOP;
END $$;

COMMIT;

-- 4) 확인용 — 아직 "행 보안이 꺼져 있거나, 공개(anon·public) 대상 규칙이 남은 표" 목록
--    실행 결과 표를 그대로 Claude Code에 붙여 주세요. 0줄이면 가장 좋습니다.
SELECT
  c.relname AS "표",
  c.relrowsecurity AS "행보안켜짐",
  COALESCE(string_agg(pol.policyname || ' [' || pol.cmd || ' → ' || array_to_string(pol.roles, ',') || ']', ' / '), '(규칙 없음)') AS "규칙"
FROM pg_class c
JOIN pg_namespace n ON n.oid = c.relnamespace
LEFT JOIN pg_policies pol ON pol.schemaname = 'public' AND pol.tablename = c.relname
WHERE n.nspname = 'public' AND c.relkind = 'r'
GROUP BY c.relname, c.relrowsecurity
HAVING NOT c.relrowsecurity
    OR bool_or(pol.roles && ARRAY['anon','public']::name[])
ORDER BY 1;

-- ── 되돌리기(문제가 생겼을 때만) ──
-- 특정 표를 예전처럼 열려면:
--   CREATE POLICY "<표>_all" ON public.<표> FOR ALL USING (true) WITH CHECK (true);
-- 함수는:
--   GRANT EXECUTE ON FUNCTION public.<함수>(...) TO anon;
