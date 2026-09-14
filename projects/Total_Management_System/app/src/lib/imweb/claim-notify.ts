/**
 * 아임웹 주문 취소·반품 웹훅 → 고객 알림톡 / 사장님 푸시 (2026-09-14)
 *
 * 흐름: /api/imweb/webhook 가 원본을 imweb_webhook_events 에 기록 → 취소·반품 이벤트면 여기서 처리 → 결과를 같은 행에 기록
 *
 * 설계 근거 = 아임웹 "테스트 보내기" 10종 실측 (노션 🛒 아임웹 주문 관련 템플릿 · 🧪 v2 카드)
 *  - 취소·반품 웹훅엔 고객 이름·전화가 없음 → TMS 주문 DB(orders) 우선, 없으면 v2 주문조회(key/secret — OAuth 토큰 불필요)
 *  - 품목은 웹훅 data.section.sectionItems(요청·처리된 품목만 — 부분취소 정확) 우선, 없으면 주문 품목
 *  - 거절 사유 필드가 없음 → 거절은 고객 알림톡 대신 사장님 직접 연락(푸시로 누락 방지)
 *  - 관리자가 직접 처리한 취소·반품(isCustomerRequest=N)은 '접수' 알림 생략 (고객이 요청하지 않음)
 *
 * 항상 발송 원칙: on/off 토글 없음. 가동 시점 = 설정의 'Make 웹훅 URL (아임웹 주문)' 입력 (미입력이면 알림톡 미발송·푸시만)
 */
import { sendNotification, type NotifyTemplate } from '@/lib/notification/make-webhook';
import { getOrder } from '@/lib/imweb/client';
import { createServiceClient } from '@/lib/supabase/server';

type ClaimAction =
  | { kind: 'notify'; template: NotifyTemplate; customerRequestOnly?: boolean }
  | { kind: 'push'; label: string };

/** 처리 대상 이벤트 — 여기 없는 이벤트(철회·반품 수거완료 등)는 기록만 */
const CLAIM_ACTIONS: Record<string, ClaimAction> = {
  ORDER_CANCEL_REQUEST: { kind: 'notify', template: 'imweb_cancel_requested', customerRequestOnly: true },
  ORDER_CANCEL_COMPLETE: { kind: 'notify', template: 'imweb_cancel_completed' },
  ORDER_RETURN_REQUEST: { kind: 'notify', template: 'imweb_return_requested', customerRequestOnly: true },
  ORDER_RETURN_COLLECTING: { kind: 'notify', template: 'imweb_return_approved' },
  ORDER_RETURN_COMPLETE: { kind: 'notify', template: 'imweb_return_completed' },
  ORDER_CANCEL_REJECT: { kind: 'push', label: '취소 거절' },
  ORDER_RETURN_REJECT: { kind: 'push', label: '반품 거절' },
};

export function isImwebClaimEvent(eventType: string | null): boolean {
  return !!eventType && Object.prototype.hasOwnProperty.call(CLAIM_ACTIONS, eventType);
}

type Obj = Record<string, unknown>;
const obj = (v: unknown): Obj | undefined => (v && typeof v === 'object' && !Array.isArray(v) ? (v as Obj) : undefined);
const str = (v: unknown): string => (v == null ? '' : String(v).trim());

/** 품목 목록 → '상품명 N개' / '상품명 N개 외 N건' (빈 목록이면 '') */
export function formatClaimProducts(items: Array<{ name: string; qty: number }>): string {
  const list = items.filter((i) => i.name);
  if (list.length === 0) return '';
  const first = `${list[0].name} ${list[0].qty > 0 ? list[0].qty : 1}개`;
  return list.length > 1 ? `${first} 외 ${list.length - 1}건` : first;
}

interface ClaimInput {
  eventId: number | null;
  eventType: string;
  orderNo: string;
  payload: unknown;
}

/**
 * 결과 outcome:
 *  sent(알림톡 발송) · pushed(거절 푸시) · send_failed · not_configured(아임웹 웹훅 URL 미입력)
 *  admin_initiated(관리자 직접 처리 — 접수 알림 생략) · duplicate(이미 처리됨) · order_not_found · no_phone
 */
export async function handleImwebClaimEvent({ eventId, eventType, orderNo, payload }: ClaimInput): Promise<Obj> {
  const action = CLAIM_ACTIONS[eventType];
  if (!action) return { outcome: 'ignored' };

  const data = obj(obj(payload)?.data);
  const section = obj(data?.section);
  const sectionNo = str(section?.orderSectionNo) || null;

  // 1) 관리자가 직접 처리한 취소·반품은 '접수' 알림 생략
  if (action.kind === 'notify' && action.customerRequestOnly) {
    const info = obj(section?.cancelInfo) || obj(section?.returnInfo);
    if (str(info?.isCustomerRequest).toUpperCase() === 'N') return { outcome: 'admin_initiated', section_no: sectionNo };
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;

  // 2) 중복 방지 — 같은 이벤트·주문·섹션이 이미 발송/푸시됐으면 재발송 안 함 (아임웹 재시도 대비)
  {
    let q = db
      .from('imweb_webhook_events')
      .select('id')
      .eq('event_type', eventType)
      .eq('order_no', orderNo)
      .in('process_result->>outcome', ['sent', 'pushed'])
      .limit(1);
    if (eventId != null) q = q.neq('id', eventId);
    if (sectionNo) q = q.eq('payload->data->section->>orderSectionNo', sectionNo);
    const { data: dup } = await q;
    if (Array.isArray(dup) && dup.length > 0) return { outcome: 'duplicate', of_event_id: dup[0].id, section_no: sectionNo };
  }

  // 3) 고객 연락처·주문 품목 — TMS 주문 DB 우선
  let name = '';
  let phone = '';
  let orderItems: Array<{ name: string; qty: number }> = [];
  let contactSource = '';
  {
    const { data: order } = await db
      .from('orders')
      .select('id, orderer_name, orderer_phone')
      .eq('imweb_order_no', orderNo)
      .maybeSingle();
    if (order) {
      name = str(order.orderer_name);
      phone = str(order.orderer_phone);
      contactSource = 'tms_orders';
      const { data: items } = await db.from('order_items').select('product_name, quantity').eq('order_id', order.id);
      orderItems = (items || []).map((i: { product_name: string | null; quantity: number | null }) => ({ name: str(i.product_name), qty: Number(i.quantity) || 1 }));
    }
  }
  if (!phone) {
    try {
      const res = await getOrder(orderNo);
      const o = res?.data;
      if (o?.orderer) {
        name = name || str(o.orderer.name);
        phone = str(o.orderer.call);
        contactSource = contactSource || 'imweb_v2';
      }
    } catch {
      /* 주문 없음(테스트 전송의 가짜 주문번호 등) */
    }
  }
  if (!contactSource) return { outcome: 'order_not_found', section_no: sectionNo };

  // 4) 품목명 — 웹훅의 요청·처리 품목 우선 (부분취소 정확), 없으면 주문 품목
  const sectionItems = Array.isArray(section?.sectionItems) ? (section!.sectionItems as unknown[]) : [];
  const fromWebhook = sectionItems.map((it) => {
    const o = obj(it);
    return { name: str(obj(o?.productInfo)?.prodName), qty: Number(o?.qty) || 1 };
  });
  const productName = formatClaimProducts(fromWebhook) || formatClaimProducts(orderItems) || '주문 상품';
  const customerName = name || '고객';

  // 5-a) 거절 → 사장님 푸시 (고객 알림톡 없음)
  if (action.kind === 'push') {
    const { sendPushToAll } = await import('@/lib/firebase/send-push');
    const r = await sendPushToAll({
      title: `⚠️ 아임웹 ${action.label}됨 — 고객 연락 필요`,
      body: `${customerName}님 · 주문 ${orderNo} · ${productName}`,
      url: '/orders',
      tag: `mamoru-imweb-${eventType.toLowerCase()}-${sectionNo || orderNo}`,
    });
    return { outcome: 'pushed', push_sent: r.sent, push_failed: r.failed, section_no: sectionNo, contact_source: contactSource };
  }

  // 5-b) 고객 알림톡
  if (!phone) return { outcome: 'no_phone', template: action.template, section_no: sectionNo, contact_source: contactSource };
  const result = await sendNotification({
    template: action.template,
    phone,
    name: customerName,
    data: {
      id: orderNo,
      order_no: orderNo,
      product_name: productName,
      event_type: eventType,
      section_no: sectionNo || '',
    },
  });
  if (result.success) {
    return { outcome: 'sent', template: action.template, product_name: productName, section_no: sectionNo, contact_source: contactSource };
  }
  const notConfigured = /webhook_imweb 미설정/.test(result.error || '');
  return {
    outcome: notConfigured ? 'not_configured' : 'send_failed',
    template: action.template,
    product_name: productName,
    section_no: sectionNo,
    contact_source: contactSource,
    error: result.error || null,
  };
}
