-- 053_imweb_banners.sql — 아임웹 메인 모달 배너/팝업 원격 관리
-- MVP: 단일 배너(id='main_modal') — 향후 복수 배너 확장 가능한 구조

CREATE TABLE IF NOT EXISTS imweb_banners (
  id TEXT PRIMARY KEY,                      -- 'main_modal' | 향후 'top_strip' 등
  enabled BOOLEAN NOT NULL DEFAULT false,
  title TEXT,
  description TEXT,
  image_url TEXT,                           -- Supabase Storage public URL
  image_path TEXT,                          -- 버킷 내 파일 경로 (삭제용)
  link_url TEXT,                            -- 배너 클릭 시 이동 URL (선택)
  starts_at TIMESTAMPTZ,                    -- NULL = 즉시 시작
  ends_at TIMESTAMPTZ,                      -- NULL = 무기한
  dismiss_cookie_hours INT NOT NULL DEFAULT 24, -- "오늘 하루 보지 않기" 시간
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_by UUID REFERENCES auth.users(id)
);

-- 초기 로우 (id='main_modal', 비활성 상태)
INSERT INTO imweb_banners (id, enabled, title)
VALUES ('main_modal', false, '메인 모달 배너')
ON CONFLICT (id) DO NOTHING;

-- RLS
ALTER TABLE imweb_banners ENABLE ROW LEVEL SECURITY;

CREATE POLICY imweb_banners_admin_all ON imweb_banners
  FOR ALL TO authenticated
  USING (true) WITH CHECK (true);

-- 공개 API(/api/imweb/banner-config)는 service_role 키로 RLS 우회

COMMENT ON TABLE imweb_banners IS '아임웹 사이트 배너/팝업 설정 (TMS에서 관리, widget.js로 고객 노출)';
COMMENT ON COLUMN imweb_banners.id IS '배너 슬롯 ID (main_modal 고정, 복수 배너 확장 시 다른 값)';
COMMENT ON COLUMN imweb_banners.enabled IS '노출 여부 토글';
COMMENT ON COLUMN imweb_banners.image_url IS 'Supabase Storage imweb-banners 버킷의 public URL';
COMMENT ON COLUMN imweb_banners.image_path IS '이미지 교체/삭제 시 참조 (버킷 내 경로)';
COMMENT ON COLUMN imweb_banners.dismiss_cookie_hours IS '고객이 닫은 후 재노출까지 대기 시간 (기본 24시간)';
