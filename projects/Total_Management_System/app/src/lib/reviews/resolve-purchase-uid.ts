/**
 * 구매 후기(type=purchase) uid 정규화 — 후기 폼 API(info/submit) 공용.
 *
 * 구매 후기 uid 의 정식 값 = 아임웹 주문이면 `orders.id`(UUID), 오프라인 판매면 `offline_sales.sale_number`(OS-…).
 * 2026-09-13 이전 `api/cron/track-delivery` 아임웹 배송완료 자동후기는 uid 에 **아임웹 주문번호(imweb_order_no)** 를 넣어 보냈다.
 *   → 이미 고객 카톡에 나간 링크가 "주문 정보를 찾을 수 없습니다" 로 막히지 않도록, 숫자 주문번호면 orders.id 로 바꿔 준다.
 * 정규화된 값으로 중복체크(source_id = uid:productNo)까지 해야 옛 링크·새 링크가 같은 후기로 묶인다.
 */

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const IMWEB_ORDER_NO_RE = /^\d{10,}$/;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export async function resolvePurchaseUid(db: any, uid: string): Promise<string> {
  if (UUID_RE.test(uid) || !IMWEB_ORDER_NO_RE.test(uid)) return uid; // UUID·판매번호(OS-) 는 그대로
  const { data } = await db.from('orders').select('id').eq('imweb_order_no', uid).maybeSingle();
  return (data?.id as string | undefined) || uid;
}
