-- 054_imweb_banner_images.sql — 아임웹 배너 슬라이드 지원 (이미지 다중화)
-- 정책: 이미지 1장=정적 / 2장+=슬라이드(5초 자동전환, 스와이프/점/화살표)
-- 최대 5장 (애플리케이션 레이어에서 검증)

-- 1. images JSONB 컬럼 추가
-- 구조: [{ url, path, link_url? }, ...]   (배열 순서 = 슬라이드 순서)
ALTER TABLE imweb_banners
  ADD COLUMN IF NOT EXISTS images JSONB NOT NULL DEFAULT '[]'::jsonb;

-- 2. 기존 image_url/image_path/link_url 데이터를 images 배열로 마이그레이션
-- (이미 images가 채워져 있지 않은 row만 변환)
UPDATE imweb_banners
SET images = jsonb_build_array(
  jsonb_build_object(
    'url',      image_url,
    'path',     COALESCE(image_path, ''),
    'link_url', COALESCE(link_url,   '')
  )
)
WHERE image_url IS NOT NULL
  AND image_url != ''
  AND (images IS NULL OR jsonb_array_length(images) = 0);

COMMENT ON COLUMN imweb_banners.images IS
  'JSONB 배열: [{url,path,link_url?}, ...]. 1개=정적, 2개+=슬라이드(5초). 최대 5장 (애플리케이션 레이어 검증).';

-- 3. 기존 image_url / image_path / link_url 컬럼은 backwards-compat 용도로 보존
-- (첫 번째 이미지의 shortcut — API 레이어에서 자동 동기화)
