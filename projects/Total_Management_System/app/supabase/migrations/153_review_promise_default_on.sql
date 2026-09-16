-- 153: 후기 요청 자동 발송을 "기본 ON" 으로 (사장님 확정 2026-09-16, C안)
--
-- 배경
--   2026-05-26 정책: "후기 약속 ✓ 받은 고객만 자동 발송" (보수적)
--   → 사장님이 매 건 체크를 눌러야 자동 발송됨. 깜빡하면 후기 요청이 통째로 누락.
--   → 아임웹 주문은 무조건 발송이라 경로마다 동작이 달랐다.
--
-- 이번 변경 (C안 = 기본값만 뒤집기)
--   컬럼 의미·UI 구조·발송 코드의 가드는 그대로 두고 **기본값만** ON 으로 바꾼다.
--   즉 "빼고 싶은 고객만 해제" 로 뒤집는다. 해제하면 종전과 100% 동일하게 동작.
--
-- 적용 대상
--   ✅ repairs(복원수리) · offline_sales(판매관리)
--   ❌ consultations — 2026-04-30 IA 정리로 리뷰 진입점이 아니다 (건드리지 않음)
--   ❌ deliveries(B2B 납품) — 후기 컬럼 자체가 없다. 사장님 지시대로 그대로 유지
--
-- 안전성
--   코드 어디에서도 INSERT 시 review_promised_at 을 명시로 넘기지 않는다(전수 grep 확인)
--   → DEFAULT 가 그대로 먹는다. 기존 행은 건드리지 않고 아래 백필로만 제한적으로 채운다.

-- ─────────────────────────────────────────────
-- 1) 신규 행: 기본 ON
-- ─────────────────────────────────────────────
ALTER TABLE repairs       ALTER COLUMN review_promised_at   SET DEFAULT now();
ALTER TABLE repairs       ALTER COLUMN review_promised_type SET DEFAULT 'repair';

ALTER TABLE offline_sales ALTER COLUMN review_promised_at   SET DEFAULT now();
ALTER TABLE offline_sales ALTER COLUMN review_promised_type SET DEFAULT 'purchase';

-- ─────────────────────────────────────────────
-- 2) 백필 — "아직 진행 중인 건"만
--    🚨 이미 배송완료/완료된 과거 건은 제외한다.
--       소급해서 채우면 다음 크론에서 옛 고객에게 후기 알림톡이 쏟아진다.
-- ─────────────────────────────────────────────

-- 복원수리: 아직 배송완료 전 + 후기 미발송 + 미작성
UPDATE repairs
SET    review_promised_at   = now(),
       review_promised_type = COALESCE(review_promised_type, 'repair')
WHERE  review_promised_at    IS NULL
  AND  review_request_sent_at IS NULL
  AND  review_submitted_at   IS NULL
  AND  status IN ('intake', 'pickup_scheduled', 'cost_notified', 'repairing',
                  'ready_to_ship', 'shipped',
                  'picked_up', 'inspecting', 'payment_confirmed');  -- 레거시 상태 포함

-- 판매관리: 아직 배송완료 전 + 미취소 + 후기 미발송 + 미작성
UPDATE offline_sales
SET    review_promised_at   = now(),
       review_promised_type = COALESCE(review_promised_type, 'purchase')
WHERE  review_promised_at   IS NULL
  AND  review_requested_at  IS NULL
  AND  review_submitted_at  IS NULL
  AND  cancelled_at         IS NULL
  AND  delivered_at         IS NULL;

-- ─────────────────────────────────────────────
-- 3) 확인용 (실행 후 눈으로 보기)
-- ─────────────────────────────────────────────
SELECT '복원수리' AS 구분,
       COUNT(*) FILTER (WHERE review_promised_at IS NOT NULL) AS "자동발송 ON",
       COUNT(*) FILTER (WHERE review_promised_at IS NULL)     AS "해제/과거"
FROM   repairs
UNION ALL
SELECT '판매관리',
       COUNT(*) FILTER (WHERE review_promised_at IS NOT NULL),
       COUNT(*) FILTER (WHERE review_promised_at IS NULL)
FROM   offline_sales;
