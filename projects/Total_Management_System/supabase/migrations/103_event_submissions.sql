-- 103: EVENT 고객 접수 (재고 전환 이벤트 등 접수형 이벤트 공통 허브)
-- 파이프라인: received(접수) → payment_noticed(입금안내) → paid(입금확인) → converted(판매전환)
-- 입금확인 시 offline_sales 자동 생성(sale_id 연결). 발송/배송/후기는 전환된 판매가 기존 인프라로 처리.
-- ※ 파일 번호는 명목상(074 이후 마이그레이션은 SQL Editor 수동 실행). 충돌 시 사장님이 번호만 조정.

create table if not exists event_submissions (
  id uuid primary key default gen_random_uuid(),
  event_number text unique not null,                 -- EV-YYYYMMDD-NNN
  customer_id uuid references customers(id),
  customer_name text not null,
  customer_phone text,
  phone_normalized text generated always as (regexp_replace(coalesce(customer_phone, ''), '\D', '', 'g')) stored,
  receive_method text not null default 'delivery',   -- 'delivery'(택배) | 'visit'(매장방문)
  postcode text,
  address1 text,
  address2 text,
  items jsonb not null default '[]'::jsonb,           -- [{product_id, product_name, category_type, spec, slicing, qty, unit_price}]
  slicing_addon integer not null default 0,           -- 슬라이싱 가공 추가비 합계
  total_amount integer not null default 0,
  status text not null default 'received',            -- received | payment_noticed | paid | converted | cancelled
  payment_noticed_at timestamptz,
  paid_at timestamptz,
  sale_id uuid references offline_sales(id),          -- 판매 전환 시 연결
  memo text,
  cancelled_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists idx_event_submissions_status on event_submissions(status);
create index if not exists idx_event_submissions_phone on event_submissions(phone_normalized);
create index if not exists idx_event_submissions_created on event_submissions(created_at desc);

create table if not exists event_history (
  id uuid primary key default gen_random_uuid(),
  event_id uuid not null references event_submissions(id) on delete cascade,
  to_status text,
  note text,
  created_at timestamptz not null default now()
);
create index if not exists idx_event_history_event on event_history(event_id);

-- updated_at 자동 갱신 (기존 테이블과 동일 패턴이 있으면 그 트리거 함수를 재사용)
create or replace function set_event_submissions_updated_at() returns trigger as $$
begin new.updated_at = now(); return new; end;
$$ language plpgsql;

drop trigger if exists trg_event_submissions_updated_at on event_submissions;
create trigger trg_event_submissions_updated_at
  before update on event_submissions
  for each row execute function set_event_submissions_updated_at();
