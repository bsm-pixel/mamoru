-- 106_customer_merge.sql — 고객 병합(merge) 인프라
-- 목적: 같은 사람이 다른 이름/전화로 중복 등록된 고객 레코드를 1개로 통합.
--       흡수된 고객은 삭제하지 않고 merged_into_id 로 soft 보존(이력 추적).

-- 1) soft 병합 마커 컬럼
alter table customers add column if not exists merged_into_id uuid references customers(id);
alter table customers add column if not exists merged_at timestamptz;
create index if not exists idx_customers_merged_into on customers(merged_into_id);

-- 2) merge_customers — 흡수 대상(p_victims)의 모든 거래를 주 고객(p_primary)으로 단일 트랜잭션 이관
--    매출/청구 문서: customer_id 이관 + 표시용 이름·전화를 주 고객으로 통일(매출·송장 한 이름)
--    접수/주문: customer_id 만 이관(원 접수자 이름은 기록 보존)
create or replace function merge_customers(p_primary uuid, p_victims uuid[])
returns json
language plpgsql
security definer
set search_path = public
as $$
declare
  v_name  text;
  v_phone text;
begin
  if p_primary = any(p_victims) then
    raise exception '주 고객은 흡수 대상에 포함될 수 없습니다';
  end if;

  select name, phone into v_name, v_phone
  from customers where id = p_primary and merged_into_id is null;
  if not found then
    raise exception '주 고객을 찾을 수 없거나 이미 병합된 고객입니다';
  end if;

  -- 매출/청구 문서 — customer_id 이관 + denormalized 이름/전화 통일
  update offline_sales   set customer_id = p_primary, customer_name = v_name, customer_phone = v_phone where customer_id = any(p_victims);
  update deliveries      set customer_id = p_primary, customer_name = v_name, customer_phone = v_phone where customer_id = any(p_victims);
  update contracts       set customer_id = p_primary, customer_name = v_name, customer_phone = v_phone where customer_id = any(p_victims);
  update manual_invoices set customer_id = p_primary, customer_name = v_name, customer_phone = v_phone where customer_id = any(p_victims);

  -- 접수/주문 — customer_id 만 이관 (접수자 본명 기록 보존)
  update repairs       set customer_id = p_primary where customer_id = any(p_victims);
  update consultations set customer_id = p_primary where customer_id = any(p_victims);
  update orders        set customer_id = p_primary where customer_id = any(p_victims);

  -- 흡수 고객 soft 보존 (목록/검색 숨김, 미수금 0 — 거래는 모두 주 고객으로 이관됨)
  update customers
    set merged_into_id = p_primary, merged_at = now(), outstanding_balance = 0
    where id = any(p_victims);

  return json_build_object('primary', p_primary, 'victims', coalesce(array_length(p_victims, 1), 0));
end;
$$;

-- 참고: 주 고객 outstanding_balance 재계산은 애플리케이션 recalcOutstanding(단일 출처)이 RPC 후 수행.
