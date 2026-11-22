-- 172: 계약서 — 고객이 링크로 폰에서 서명 + 이니셜 각인 (2026-11-22)
--
-- 온라인(톡상담) 고객에게도 계약서를 받기 위해:
--   · sign_token     : 고객 서명 링크의 열쇠(추측 불가 uuid). 링크를 만들 때만 생긴다.
--   · signed_via     : 어디서 서명했는지 'store'(매장·출장에서 직원 화면) | 'link'(고객 폰 링크)
--   · terms_agreed_at: 고객이 유의사항(교환·각인 조건)에 동의한 시각
--   · engraving*     : 이니셜 각인. 문구를 나중에 받기로 하면 engraving_status = 'pending',
--                      고객이 같은 링크로 보내거나 직원이 입력하면 'received'
-- 기존 칸은 바꾸지 않는다. 여러 번 실행해도 안전(IF NOT EXISTS).

ALTER TABLE contracts
  ADD COLUMN IF NOT EXISTS sign_token uuid,
  ADD COLUMN IF NOT EXISTS sign_link_at timestamptz,
  ADD COLUMN IF NOT EXISTS signed_via text,
  ADD COLUMN IF NOT EXISTS terms_agreed_at timestamptz,
  ADD COLUMN IF NOT EXISTS engraving boolean NOT NULL DEFAULT false,
  ADD COLUMN IF NOT EXISTS engraving_text text,
  ADD COLUMN IF NOT EXISTS engraving_status text,
  ADD COLUMN IF NOT EXISTS engraving_at timestamptz;

CREATE UNIQUE INDEX IF NOT EXISTS uq_contracts_sign_token ON contracts(sign_token) WHERE sign_token IS NOT NULL;

COMMENT ON COLUMN contracts.sign_token IS '고객 서명 링크 열쇠 (page.mamoru.kr/projects/contract/page_sign.html?t=)';
COMMENT ON COLUMN contracts.engraving_status IS 'NULL=각인 없음 · pending=문구 나중에 받기 · received=문구 받음';
