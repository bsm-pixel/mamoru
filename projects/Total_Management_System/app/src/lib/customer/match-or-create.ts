/**
 * 고객 자동 매칭/생성 helper — phone 기반 SSOT 강화 (2026-04-30)
 *
 * 사용처:
 *   - /api/consultation/public/submit (고객 폼)
 *   - /api/consultation/admin-create (관리자 수기 등록)
 *   - /api/repair/public/submit (복원수리 고객 폼)
 *   - /api/repair POST (관리자 수기 등록)
 *
 * 동작:
 *   1. phone을 정규화(digits-only) → phone_normalized
 *   2. phoneNorm 비어있으면 null 반환 (호출 측에서 customer_id NULL로 INSERT)
 *   3. customers.phone_normalized = phoneNorm 검색 (가장 오래된 customer 매칭, deterministic)
 *   4. 매칭됨 → 기존 customerId 반환
 *   5. 매칭 X → 신규 INSERT → 신규 customerId 반환
 *
 * Edge case:
 *   - 동명이인 (같은 phone, 다른 이름): 첫 매칭 사용 (운영 현실에서 드뭄, 사장님이 사후 분리 가능)
 *   - INSERT 실패: 에러 로그 + null 반환 (호출 측에서 customer_id 없이 INSERT 진행)
 */

export interface MatchOrCreateInput {
  phone: string;
  name: string;
  source: 'consultation' | 'as' | 'manual' | 'event';
  extra?: {
    addressRoad?: string | null;
    addressDetail?: string | null;
    postcode?: string | null;
    customerType?: string;
    activityName?: string | null;   // 102: 활동명(매장 사용 이름)
    position?: string | null;        // 102: 직급
  };
}

export interface MatchOrCreateResult {
  customerId: string | null;
  isNew: boolean;
}

export async function matchOrCreateCustomer(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  db: any,
  input: MatchOrCreateInput,
): Promise<MatchOrCreateResult> {
  const phoneNorm = (input.phone || '').replace(/\D/g, '');
  if (!phoneNorm) return { customerId: null, isNew: false };

  // 1) 기존 매칭 시도 — 활성(비병합) 고객 우선, 가장 오래된 것 (deterministic)
  //    병합으로 숨긴 고객을 잡으면 새 접수가 숨긴 레코드에 붙어버리므로 merged_into_id IS NULL 만 매칭
  let existing: Array<{ id: string }> | null = null;
  const { data: active, error: searchErr } = await db
    .from('customers')
    .select('id')
    .eq('phone_normalized', phoneNorm)
    .is('merged_into_id', null)
    .order('created_at', { ascending: true })
    .limit(1);

  if (searchErr) {
    console.error('[matchOrCreateCustomer] 검색 실패:', searchErr);
    return { customerId: null, isNew: false };
  }
  existing = active;

  // 활성 고객이 없고 병합된 레코드만 있으면 → 그 주 고객(merged_into_id)으로 연결
  if (!existing || existing.length === 0) {
    const { data: merged } = await db
      .from('customers')
      .select('merged_into_id')
      .eq('phone_normalized', phoneNorm)
      .not('merged_into_id', 'is', null)
      .order('created_at', { ascending: true })
      .limit(1);
    if (merged && merged.length > 0 && merged[0].merged_into_id) {
      existing = [{ id: merged[0].merged_into_id as string }];
    }
  }

  const activityName = input.extra?.activityName?.trim() || null;
  const position = input.extra?.position?.trim() || null;

  if (existing && existing.length > 0) {
    const id = existing[0].id as string;
    const patch: Record<string, unknown> = {};
    // 활동명/직급은 비어있을 때만 채움 (수기 큐레이션 값 덮어쓰기 방지)
    if (activityName || position) {
      const { data: cur } = await db.from('customers').select('activity_name, position').eq('id', id).single();
      if (activityName && !cur?.activity_name) patch.activity_name = activityName;
      if (position && !cur?.position) patch.position = position;
    }
    // 주소: 접수(택배)에 도로명 주소가 있으면 전체를 접수값으로 최신화 (정책 B — 최신 배송지 반영)
    if (input.extra?.addressRoad) {
      patch.address_road = input.extra.addressRoad;
      patch.address_detail = input.extra.addressDetail || null;
      patch.postcode = input.extra.postcode || null;
    }
    if (Object.keys(patch).length > 0) await db.from('customers').update(patch).eq('id', id);
    return { customerId: id, isNew: false };
  }

  // 2) 신규 INSERT — phone_normalized는 DB trigger로 자동 채워짐
  const insertData: Record<string, unknown> = {
    name: input.name?.trim() || '미기입',
    phone: input.phone,
    source: input.source,
    customer_type: input.extra?.customerType || 'retail',
  };
  if (input.extra?.addressRoad) insertData.address_road = input.extra.addressRoad;
  if (input.extra?.addressDetail) insertData.address_detail = input.extra.addressDetail;
  if (input.extra?.postcode) insertData.postcode = input.extra.postcode;
  if (activityName) insertData.activity_name = activityName;
  if (position) insertData.position = position;

  const { data: created, error: insertErr } = await db
    .from('customers')
    .insert(insertData)
    .select('id')
    .single();

  if (insertErr || !created) {
    console.error('[matchOrCreateCustomer] INSERT 실패:', insertErr);
    return { customerId: null, isNew: false };
  }

  return { customerId: created.id as string, isNew: true };
}
