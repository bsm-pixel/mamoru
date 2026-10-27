/**
 * 165 제품·복원수리 함께 출고 알림톡 (combined_shipped) — 2026-10-27
 *
 * 같은 송장으로 판매 제품과 복원수리 가위가 한 상자로 나가면, 고객에게 출고 안내를 **1통**만 보낸다
 *   (전엔 as_shipped + sales_shipped 2통 — 같은 송장번호가 두 번 와서 헷갈렸다).
 *
 * 진입점: api/cron/track-delivery [2-A] 복원수리 집하 감지.
 *   ① claimSaleForCombined — 루프 안에서 **동기로** 같은 송장의 판매를 선점(shipped_at CAS).
 *      after() 안에서 하면 같은 사이클의 [3-A] 판매 집하가 먼저 잡아 sales_shipped 가 따로 나간다.
 *   ② sendCombinedShipped — after() 안에서 발송. 실패하면 기존 2통으로 되돌려 보낸다(안내 누락 금지).
 *
 * 가동 게이트: 설정 notifications.webhook_field_repair 가 비어 있으면 ①이 null → 지금처럼 2통.
 *   (솔라피 검수·Make 분기 전에 1통으로 바꿔 버리면 고객이 출고 안내를 한 통도 못 받는다)
 */
import { sendNotification, isFieldRepairWebhookReady } from '@/lib/notification/make-webhook';
import { sendSalesShippedNotification } from '@/lib/notification/sales-shipped';
import { isB2BCustomerType } from '@/lib/sales/customer-type';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface CombinedSale {
  id: string;
  sale_number: string | null;
  customer_name: string | null;
  customer_phone: string | null;
  customer_type: string | null;
  courier_name: string | null;
  invoice_number: string;
}

/** 같은 송장의 B2C 판매(미출고)를 출고 처리로 선점. 함께 출고 대상이 아니면 null */
export async function claimSaleForCombined(db: Db, invoiceNumber: string, pickedUpAt: string): Promise<CombinedSale | null> {
  if (!(await isFieldRepairWebhookReady())) return null;
  const { data: sales } = await db
    .from('offline_sales')
    .select('id, sale_number, customer_name, customer_phone, customer_type, courier_name, invoice_number')
    .eq('invoice_number', invoiceNumber)
    .is('shipped_at', null)
    .is('cancelled_at', null)
    .limit(1);
  const sale = (sales || [])[0] as CombinedSale | undefined;
  if (!sale || !sale.customer_phone || isB2BCustomerType(sale.customer_type)) return null;

  // [3-A] 와 같은 CAS — shipped_at 이 아직 비어 있을 때만 (크론 중복·수동 버튼과 경합해도 1회)
  const { data: claimed } = await db
    .from('offline_sales')
    .update({ shipped_at: pickedUpAt, shipped_source: 'alps_pickup' })
    .eq('id', sale.id)
    .is('shipped_at', null)
    .select('id');
  return claimed && claimed.length > 0 ? sale : null;
}

/** 제품명만 (복원수리 RS 줄 제외) — "블런트 5.5 외 1건" */
async function productSummary(db: Db, saleId: string): Promise<string> {
  const { data: items } = await db
    .from('offline_sale_items')
    .select('product_name, quantity, category')
    .eq('sale_id', saleId);
  const goods = ((items || []) as Array<{ product_name: string; quantity: number; category: string | null }>)
    .filter((i) => i.category !== 'RS');
  if (goods.length === 0) return '마모루 제품';
  const first = goods[0].quantity > 1 ? `${goods[0].product_name} ${goods[0].quantity}개` : goods[0].product_name;
  return goods.length > 1 ? `${first} 외 ${goods.length - 1}건` : first;
}

/**
 * 함께 출고 1통 발송. 실패 시 기존 2통(as_shipped + sales_shipped)으로 되돌린다.
 * 반환 combined=true 면 1통으로 나갔다.
 */
export async function sendCombinedShipped(
  db: Db,
  repair: { as_id: string; name: string; phone: string; qty: number; courier_name: string | null; invoice_number: string },
  sale: CombinedSale,
): Promise<{ combined: boolean }> {
  const courier = repair.courier_name || '롯데택배';
  const r = await sendNotification({
    template: 'combined_shipped',
    phone: repair.phone,
    name: repair.name,
    data: {
      id: repair.as_id,
      as_id: repair.as_id,
      as_uid: repair.as_id,              // [수리내역서 보기] 버튼
      qty: String(repair.qty),
      product_name: await productSummary(db, sale.id),
      tracking: repair.invoice_number,   // [배송 조회] 버튼 · 송장번호
      courier,
    },
  });
  if (r.success) {
    await db.from('offline_sales').update({ shipped_notified_at: new Date().toISOString() }).eq('id', sale.id);
    return { combined: true };
  }

  // 되돌림 — 고객이 출고 안내를 못 받는 일은 없어야 한다
  console.error(`[combined-shipped] ${repair.as_id} 1통 발송 실패 → 2통으로 되돌림:`, r.error);
  await sendNotification({
    template: 'as_shipped',
    phone: repair.phone,
    name: repair.name,
    data: { id: repair.as_id, as_id: repair.as_id, as_uid: repair.as_id, courier, tracking: repair.invoice_number },
  });
  const s = await sendSalesShippedNotification(db, {
    id: sale.id,
    saleNumber: sale.sale_number,
    invoiceNumber: sale.invoice_number,
    customerName: sale.customer_name || '고객',
    customerPhone: sale.customer_phone,
    customerType: sale.customer_type,
    courierName: sale.courier_name,
  });
  if (s.sent) await db.from('offline_sales').update({ shipped_notified_at: new Date().toISOString() }).eq('id', sale.id);
  return { combined: false };
}
