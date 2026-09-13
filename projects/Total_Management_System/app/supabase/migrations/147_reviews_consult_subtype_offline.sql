-- 147: 상담 리뷰 subtype 레거시 'offline' → 'store_visit' (칩 '상담·오프라인' → '상담·직접방문')
-- 배경(2026-09-13): 146 에 이어, 4분류(2026-07-17) 이전 판매채널 'offline' 으로 저장된 상담 리뷰도
--   직접방문으로 통일(사장님 결정). 코드도 saleChannelToConsultSubtype 에서 offline→store_visit 로 정규화.
-- 예상 1건: RV-20260625-001. 멱등(재실행 0건).
-- ✅ 2026-09-13 적용 완료(Claude, 동일 조건 REST PATCH 1건) — 수동 재실행 불필요(해도 0건).
UPDATE reviews
   SET subtype = 'store_visit'
 WHERE type = 'consult'
   AND subtype = 'offline';

-- 확인(0건이어야 함): SELECT count(*) FROM reviews WHERE type = 'consult' AND subtype = 'offline';
