-- ╔══════════════════════════════════════════════════════════════════════╗
-- ║ 146_reviews_consult_subtype_normalize — 상담 리뷰 subtype 어휘 정규화 ║
-- ╚══════════════════════════════════════════════════════════════════════╝
-- 배경(2026-09-13):
--   판매건(OS-) 상담 후기를 리뷰 약속 칩 없이 요청하면 /api/reviews/submit 이
--   판매채널(sale_channel) 원시값 'store' 를 subtype 으로 저장했다.
--   → 관리자 '상담·매장' / 고객 페이지 '상담'(라벨 없음)으로 갈라져 표시.
--   코드 fix: lib/reviews/consult-subtype.ts saleChannelToConsultSubtype (신규 저장분부터 정규화).
--   이 SQL = 이미 저장된 과거분 교정. 멱등(재실행해도 0건).
--
-- 실행 전 확인(예상 3건: RV-20260909-001 / RV-20260911-001 / RV-20260913-001):
--   SELECT review_id, subtype, source_id FROM reviews
--    WHERE type = 'consult' AND subtype IN ('store','field','talk');

UPDATE reviews
   SET subtype = CASE subtype
                   WHEN 'store' THEN 'store_visit'
                   WHEN 'field' THEN 'field_request'
                   WHEN 'talk'  THEN 'talk_consult'
                 END
 WHERE type = 'consult'
   AND subtype IN ('store', 'field', 'talk');

-- 실행 후 확인(0건이어야 함):
--   SELECT count(*) FROM reviews WHERE type = 'consult' AND subtype IN ('store','field','talk');
