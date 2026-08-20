-- 126: 직접수령(대면 픽업) 마커
-- 배송완료(delivered) 중에서도 "매장 대면 픽업(아임웹 쿠폰 결제 등)"을 구분하기 위한 플래그.
-- 라이프사이클은 delivered 그대로 두고(탭/리뷰/집계 재사용), 배지 표시 + 픽업 매출 분리 집계에 사용.
ALTER TABLE orders ADD COLUMN IF NOT EXISTS is_pickup boolean DEFAULT false;
