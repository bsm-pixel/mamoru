-- 138_tidy_exchange_return_label.sql — 소급 교환건 "(교환 전)" 라벨 정리 (2026-08-27)
-- 마이그135가 product_name 에 'A2-65FS (교환 전)' 처럼 박아 상태(완료)와 헷갈림.
-- "(교환 전)" 접미사만 제거 → 구제품명만 남김(교환 여부는 return_type 칩 + new_product 카드로 표시).
-- 안전: 해당 패턴만 대상, 재실행해도 무해(이미 제거된 건 매칭 안 됨).

UPDATE returns
   SET product_name = btrim(replace(product_name, '(교환 전)', ''))
 WHERE product_name LIKE '%(교환 전)%';
