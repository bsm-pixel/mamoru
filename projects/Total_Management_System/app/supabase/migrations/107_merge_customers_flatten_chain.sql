-- 107_merge_customers_flatten_chain.sql — 병합 체인 평탄화
-- 문제: A를 B로 병합(A.merged_into_id=B) 후, 다시 B를 C로 병합하면
--       A.merged_into_id 는 여전히 B(이미 병합된 고객)를 가리켜 A→B→C 체인 발생.
--       표시(숨김)는 정상이나 merged_into_id 추적이 한 번에 안 끝남.
-- 해결: 흡수 시, "흡수 대상(p_victims)을 가리키던 고객들"도 주 고객(p_primary)으로 재지정.

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

  -- 체인 평탄화: 흡수 대상을 가리키던 과거 병합 고객들도 주 고객으로 재지정 (A→B→C ⇒ A→C)
  update customers set merged_into_id = p_primary
    where merged_into_id = any(p_victims);

  -- 흡수 고객 soft 보존 (목록/검색 숨김, 미수금 0 — 거래는 모두 주 고객으로 이관됨)
  update customers
    set merged_into_id = p_primary, merged_at = now(), outstanding_balance = 0
    where id = any(p_victims);

  return json_build_object('primary', p_primary, 'victims', coalesce(array_length(p_victims, 1), 0));
end;
$$;
