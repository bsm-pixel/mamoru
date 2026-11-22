import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';

/**
 * 서명된 계약서에 TMS 고객을 붙인다 (2026-11-23, t-261123-sign-ux)
 *  - 상담자를 불러오지 않고 빈 계약서로 쓴 고객도 서명하면 고객이 생긴다 → [판매 입력]이 고객이 채워진 채로 열린다.
 *  - 같은 번호 고객이 있으면 그 고객에 붙인다(matchOrCreateCustomer — 병합된 고객은 주 고객으로).
 *  - 저장된 고객 주소는 바꾸지 않는다(비어 있을 때만 채움) — 출장 주소 수정과 같은 원칙.
 *  - 연락처가 없으면 아무것도 하지 않는다. 실패해도 계약서 저장은 그대로(로그만).
 */
export async function linkContractCustomer(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  db: any,
  c: { id: string; customer_id?: string | null; customer_name?: string | null; customer_phone?: string | null },
  addr?: { road?: string | null; detail?: string | null; postcode?: string | null },
): Promise<string | null> {
  if (c.customer_id) return c.customer_id;
  if (!(c.customer_phone || '').replace(/\D/g, '')) return null;
  try {
    const { customerId } = await matchOrCreateCustomer(db, {
      phone: c.customer_phone!,
      name: c.customer_name || '',
      source: 'manual',
      extra: { addressRoad: addr?.road || null, addressDetail: addr?.detail || null, postcode: addr?.postcode || null, keepExistingAddress: true },
    });
    if (customerId) await db.from('contracts').update({ customer_id: customerId }).eq('id', c.id).is('customer_id', null);
    return customerId;
  } catch (e) {
    console.error('[contract] 고객 연결 실패', e);
    return null;
  }
}
