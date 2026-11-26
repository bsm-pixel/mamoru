-- 173: 복원수리 「콘텐츠 후보」 표시 (작업실 19번, 2026-11-26)
-- 입고 점검 때 콘텐츠로 쓸 만한 건을 표시하고, 작업 전·후 촬영 여부와 고객 노출 동의를 기록한다.
-- 기존 자료는 바꾸지 않는다(새 칸은 모두 기본값 = 후보 아님 · 촬영 안 함 · 동의 미확인).

alter table repairs
  add column if not exists content_candidate   boolean not null default false, -- 콘텐츠 후보로 표시
  add column if not exists content_symptom     text,                           -- 증상 한 줄 (예: 날 끝 이빨 빠짐)
  add column if not exists content_before_shot boolean not null default false, -- 작업 전 촬영 확인
  add column if not exists content_after_shot  boolean not null default false, -- 작업 후 촬영 확인
  add column if not exists content_consent     text;                           -- 고객 노출 동의: 'yes' | 'no' | null(미확인)

alter table repairs drop constraint if exists repairs_content_consent_check;
alter table repairs add constraint repairs_content_consent_check
  check (content_consent is null or content_consent in ('yes', 'no'));

-- 「콘텐츠 후보」 목록용 (후보인 건만 색인)
create index if not exists repairs_content_candidate_idx
  on repairs (received_at desc) where content_candidate;
