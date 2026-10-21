/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * 롯데 반품 수거접수 — 단일 출처 (2026-10-21 분리)
 *   ustRtgSctCd='02', 고객집 → 마모루 회수. 롯데 IS팀 회신(2026-08-27): 02=반품, 양식은 출고와 동일,
 *   orglInvNo=원송장(선택), **취소 API 미지원** → 접수 후 취소는 ALPS 화면에서 수동.
 *
 * 호출처
 *   · POST /api/returns/[id]/pickup — 반품·교환 메뉴 [롯데 반품 수거접수] 버튼
 *   · PATCH /api/deliveries/[id] action='return' — 납품 반품 「롯데 반품수거 접수」 선택 시 접수와 동시에(164)
 */
import { getNextInvoice, bookReturnPickup } from '@/lib/lotte/alps-client';

export type BookReturnPickupResult =
  | { ok: true; invoiceNumber: string; row: any | null; warning?: string }
  | { ok: false; status: number; error: string };

export async function bookReturnPickupFor(db: any, returnId: string): Promise<BookReturnPickupResult> {
  const { data: r } = await db.from('returns').select('*').eq('id', returnId).single();
  if (!r) return { ok: false, status: 404, error: '반품 건을 찾을 수 없습니다' };
  if (r.pickup_invoice_number) return { ok: false, status: 400, error: '이미 반품 수거접수가 되어 있습니다' };
  if (r.status === 'cancelled' || r.status === 'completed') {
    return { ok: false, status: 400, error: '취소·완료된 건은 수거접수할 수 없습니다' };
  }

  // 고객(수거지) 주소 — 고객 정보에서 조회(교환 출고와 동일 패턴)
  let custName = r.name as string | null;
  let custTel = r.phone as string | null;
  let postcode: string | null = null;
  let addrRoad: string | null = null;
  let addrDetail: string | null = null;
  if (r.customer_id) {
    const { data: c } = await db.from('customers').select('name, phone, postcode, address_road, address_detail').eq('id', r.customer_id).single();
    if (c) {
      custName = custName || c.name;
      custTel = custTel || c.phone;
      postcode = c.postcode; addrRoad = c.address_road; addrDetail = c.address_detail;
    }
  }
  // 164: 납품 반품은 접수 때 그 납품의 배송지를 반품 건에 저장해 둔다 → 그 주소를 우선 쓴다(교환 출고 ship 과 같은 폴백)
  if (r.delivery_id && r.postcode && r.address) {
    postcode = r.postcode; addrRoad = r.address; addrDetail = r.address_detail;
  }
  if (!postcode || !addrRoad) {
    return { ok: false, status: 400, error: '고객 주소(우편번호+도로명)가 없습니다. 고객 정보에서 보강 후 다시 시도해주세요.' };
  }
  if (!custTel) return { ok: false, status: 400, error: '고객 연락처가 없습니다.' };

  // 원송장 참조 — 교환 건은 교환 출고 송장, 납품 반품은 그 납품의 송장(164)
  let orgInvoice: string | undefined = r.exchange_out_invoice_number || undefined;
  if (!orgInvoice && r.delivery_id) {
    const { data: d } = await db.from('deliveries').select('tracking_number').eq('id', r.delivery_id).single();
    orgInvoice = d?.tracking_number || undefined;
  }

  const goodsName = `${r.product_name || '반품 상품'}${r.serial_number ? ` (${r.serial_number})` : ''} 회수`.slice(0, 50);
  const fullAddr = [addrRoad, addrDetail].filter(Boolean).join(' ');

  // 반품 수거접수 — 보내는분(수거지)=고객, 받는분=마모루, ustRtgSctCd='02'. 원송장은 있으면 참조(선택)
  const { invoiceNumber } = await getNextInvoice();
  const result = await bookReturnPickup({
    invoiceNumber,
    pickupName: custName || '고객',
    pickupTel: custTel,
    pickupZip: postcode,
    pickupAddr: fullAddr,
    goodsName,
    deliveryMessage: '반품 수거',
    orgInvoiceNumber: orgInvoice,  // 원송장(있으면, 선택)
  });
  if (!result.success) return { ok: false, status: 502, error: `롯데 반품 수거접수 실패: ${result.error}` };

  // 송장 저장 + 상태 '수거예약'으로 전이
  const nowIso = new Date().toISOString();
  const { data, error } = await db.from('returns').update({
    pickup_invoice_number: invoiceNumber,
    pickup_courier_name: '롯데택배',
    pickup_booked_at: nowIso,
    status: 'pickup_scheduled',
    pickup_scheduled_at: nowIso,
  }).eq('id', returnId).select().single();
  if (error) {
    console.error('[returns pickup] DB 저장 실패(ALPS는 성공):', invoiceNumber, error);
    return { ok: true, invoiceNumber, row: null, warning: 'DB 저장 실패 — 송장은 발급됨(ALPS 확인)' };
  }
  return { ok: true, invoiceNumber, row: data };
}
