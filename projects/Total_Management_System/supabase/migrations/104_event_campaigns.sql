-- 104: EVENT 캠페인 — 이벤트별(재고전환/한정판매/공동구매/체험단 등) 분리 + 허브 카드
create table if not exists event_campaigns (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  type text not null default 'stock_clearance',  -- stock_clearance|limited|group_buy|tester|trade_in|other
  status text not null default 'active',          -- active | ended
  is_default boolean not null default false,
  starts_at date,
  ends_at date,
  memo text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table event_submissions add column if not exists campaign_id uuid references event_campaigns(id);
create index if not exists idx_event_submissions_campaign on event_submissions(campaign_id);

-- 기본 캠페인 시드 (재고 전환 EVENT 1탄) + 기존 접수 backfill
insert into event_campaigns (name, type, status, is_default)
  select '재고 전환 EVENT', 'stock_clearance', 'active', true
  where not exists (select 1 from event_campaigns where is_default = true);

update event_submissions
  set campaign_id = (select id from event_campaigns where is_default = true order by created_at limit 1)
  where campaign_id is null;

create or replace function set_event_campaigns_updated_at() returns trigger as $$
begin new.updated_at = now(); return new; end;
$$ language plpgsql;
drop trigger if exists trg_event_campaigns_updated_at on event_campaigns;
create trigger trg_event_campaigns_updated_at
  before update on event_campaigns
  for each row execute function set_event_campaigns_updated_at();
