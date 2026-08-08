-- ═══════════════════════════════════════════════════════════════════
-- 청첩장 참석여부(RSVP) 저장 테이블
-- Supabase SQL Editor 에서 "1회" 실행하세요.
-- 프로젝트: zuqabeeurtwpzmkdsynx (TMS와 동일 프로젝트 / 전용 테이블 wedding_rsvp)
-- ═══════════════════════════════════════════════════════════════════

create table if not exists public.wedding_rsvp (
  id          bigint generated always as identity primary key,
  created_at  timestamptz not null default now(),
  attending   text not null check (attending in ('참석','미참석')),
  name        text not null check (char_length(name) between 1 and 40),
  phone       text        check (char_length(phone) <= 20),
  headcount   int         check (headcount between 1 and 20),
  meal        text        check (meal in ('예정','안함')),
  message     text        check (char_length(message) <= 500)
);

alter table public.wedding_rsvp enable row level security;

-- 익명(anon)은 INSERT만 허용(응답 작성). SELECT/UPDATE/DELETE 정책 없음 → 남의 응답 조회·수정·삭제 불가.
drop policy if exists "wedding_rsvp anon insert" on public.wedding_rsvp;
create policy "wedding_rsvp anon insert"
  on public.wedding_rsvp
  for insert
  to anon
  with check (true);

-- ── 응답 확인 방법 ──
-- Supabase 대시보드 → Table editor → wedding_rsvp
-- 또는 SQL:  select * from public.wedding_rsvp order by created_at desc;
