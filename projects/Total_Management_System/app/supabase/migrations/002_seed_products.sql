-- ============================================
-- MAMORU TMS — 제품 시드 데이터 (products.js 기반)
-- ============================================

INSERT INTO products (sku, name, category, price, image_url, tags, is_active) VALUES
  ('BL-001', '마모루 소프트 블런트 5.5', 'BL', 350000, NULL, '{"type":["BL"],"stage":["DE","IN"],"feel":["FEEL_SOFT"],"style":["St_BACK","St_NONE"],"habit":["HAB_WET","HAB_NONE"],"gender":["FM","M"]}', true),
  ('BL-002', '마모루 라이트 블런트 5.5', 'BL', 280000, NULL, '{"type":["BL"],"stage":["CE","IN"],"feel":["FEEL_SOFT","FEEL_NONE"],"style":["St_GO","St_NONE"],"habit":["HAB_WET","HAB_DRY","HAB_NONE"],"gender":["FM","M"]}', true),
  ('BL-003', '마모루 파워 블런트 6.0', 'BL', 380000, NULL, '{"type":["BL"],"stage":["DE"],"feel":["FEEL_POWER"],"style":["St_GO"],"habit":["HAB_DRY","HAB_NONE"],"gender":["M"]}', true),
  ('BL-004', '마모루 클래식 블런트 5.5', 'BL', 320000, NULL, '{"type":["BL"],"stage":["DE","IN"],"feel":["FEEL_POWER","FEEL_NONE"],"style":["St_GO","St_BACK"],"habit":["HAB_WET","HAB_DRY","HAB_NONE"],"gender":["FM","M"]}', true),
  ('TH-001', '마모루 프리미엄 틴닝 25%', 'TH', 180000, NULL, '{"type":["TH"],"stage":["DE","IN","CE"],"thRatio":["TH_25"],"gender":["FM","M"]}', true),
  ('TH-002', '마모루 스탠다드 틴닝 25%', 'TH', 120000, NULL, '{"type":["TH"],"stage":["CE","IN"],"thRatio":["TH_25"],"gender":["FM","M"]}', true),
  ('TH-003', '마모루 정밀 틴닝 15%', 'TH', 160000, NULL, '{"type":["TH"],"stage":["DE"],"thRatio":["TH_15"],"gender":["FM","M"]}', true),
  ('TH-004', '마모루 쾌속 틴닝 35%', 'TH', 150000, NULL, '{"type":["TH"],"stage":["DE","IN"],"thRatio":["TH_35"],"gender":["FM","M"]}', true),
  ('LO-001', '마모루 장가위 7.0 (블런트 겸용)', 'LO', 420000, NULL, '{"type":["LO"],"stage":["DE"],"loUse":["LO_BL"],"feel":["FEEL_SOFT","FEEL_POWER"],"gender":["FM","M"]}', true),
  ('LO-002', '마모루 싱글링 전용 장가위 7.0', 'LO', 380000, NULL, '{"type":["LO"],"stage":["DE"],"loUse":["LO_SING"],"gender":["FM","M"]}', true),
  ('SL-001', '마모루 슬라이싱 가위 6.0', 'SL', 280000, NULL, '{"type":["SL"],"stage":["DE","IN"],"slWhy":["SL_NEW"],"gender":["FM","M"]}', true),
  ('SL-002', '마모루 프로 슬라이싱 6.0', 'SL', 350000, NULL, '{"type":["SL"],"stage":["DE"],"slWhy":["SL_SAME"],"slSameWhy":["SL_SAME_WHY_UNCOM","SL_SAME_WHY_UNCOM1"],"gender":["FM","M"]}', true)
ON CONFLICT (sku) DO NOTHING;
