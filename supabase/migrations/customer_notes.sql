-- ═══════════════════════════════════════════════════════════════════
-- 고객 메모 타임라인 (상담사 관점: 특징/불편/요구 등 날짜별 기록)
-- Supabase SQL Editor 에서 1회 실행
-- ═══════════════════════════════════════════════════════════════════

create table if not exists public.customer_notes (
  id          uuid primary key default gen_random_uuid(),
  customer_id uuid not null references public.customers(id) on delete cascade,
  body        text not null check (char_length(body) between 1 and 2000),
  category    text check (category in ('특징','불편','요구','기타')),
  created_at  timestamptz not null default now(),
  created_by  text,                 -- 작성자(로그인 이메일)
  deleted_at  timestamptz           -- soft delete (이력 보존)
);

-- 고객별 최신순 조회 인덱스 (삭제 안 된 것만)
create index if not exists idx_customer_notes_customer
  on public.customer_notes (customer_id, created_at desc)
  where deleted_at is null;

alter table public.customer_notes enable row level security;

-- 내부 관리도구: 로그인(authenticated) 사용자 전체 허용
drop policy if exists "customer_notes authenticated all" on public.customer_notes;
create policy "customer_notes authenticated all"
  on public.customer_notes for all
  to authenticated using (true) with check (true);
