/**
 * 아임웹 → Supabase 주문 동기화 로직
 * 증분 동기화: 마지막 synced_at 이후 변경분만 가져옴
 */

import { getOrders, getProdOrders } from './client';
import type { ImwebOrder, ImwebProdOrder } from './types';
import type { OrderStatus } from '@/lib/supabase/types';
import { createServiceClient } from '@/lib/supabase/server';
import { sendNotification } from '@/lib/notification/make-webhook';
import { subDays } from 'date-fns';

/** 아임웹 prod-order status → TMS status 매핑 */
function mapImwebStatus(status: string): OrderStatus {
  const map: Record<string, OrderStatus> = {
    BEFORE_DEPOSIT: 'pay_wait',
    PAY_COMPLETE: 'pay_done',
    PREPARE: 'preparing',
    DELIVERY: 'shipping',
    DELIVERY_COMPLETE: 'delivered',
    CONFIRM: 'confirmed',
    CANCEL: 'cancelled',
    REFUND_REQUEST: 'refund_request',
    REFUND: 'refunded',
  };
  return map[status] || 'pay_done';
}

/** 동기화 실행 */
export async function syncOrders(): Promise<{
  success: boolean;
  synced: number;
  errors: string[];
}> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  const errors: string[] = [];
  let totalSynced = 0;

  // 마지막 동기화 시점 조회
  const { data: lastSync } = await supabase
    .from('sync_log')
    .select('completed_at')
    .eq('sync_type', 'imweb_orders')
    .eq('status', 'completed')
    .order('completed_at', { ascending: false })
    .limit(1)
    .single();

  const fromTs = lastSync?.completed_at
    ? Math.floor(new Date(lastSync.completed_at).getTime() / 1000)
    : Math.floor(subDays(new Date(), 30).getTime() / 1000);

  const toTs = Math.floor(Date.now() / 1000);

  // 동기화 로그 시작
  const { data: logEntry } = await supabase
    .from('sync_log')
    .insert({
      sync_type: 'imweb_orders',
      status: 'running',
      records_synced: 0,
      started_at: new Date().toISOString(),
    })
    .select()
    .single();

  try {
    let page = 1;
    let hasMore = true;

    while (hasMore) {
      const res = await getOrders({
        order_date_from: fromTs,
        order_date_to: toTs,
        page,
        limit: 50,
      });

      const orders = res.data?.list || [];
      if (orders.length === 0) {
        hasMore = false;
        break;
      }

      for (const order of orders) {
        try {
          // 품목(prod-orders) 조회
          const prodRes = await getProdOrders(order.order_no);
          const prodOrders = prodRes.data || [];
          await upsertOrder(supabase, order, prodOrders);
          totalSynced++;
        } catch (err) {
          errors.push(`주문 ${order.order_no}: ${err}`);
        }
      }

      const pageSize = res.data?.pagenation?.pagesize || 50;
      if (orders.length < pageSize) {
        hasMore = false;
      } else {
        page++;
      }
    }

    // 동기화 로그 완료
    if (logEntry) {
      await supabase
        .from('sync_log')
        .update({
          status: 'completed',
          records_synced: totalSynced,
          error_message: errors.length > 0 ? errors.join('; ') : null,
          completed_at: new Date().toISOString(),
        })
        .eq('id', logEntry.id);
    }

    return { success: true, synced: totalSynced, errors };
  } catch (err) {
    if (logEntry) {
      await supabase
        .from('sync_log')
        .update({
          status: 'failed',
          error_message: String(err),
          completed_at: new Date().toISOString(),
        })
        .eq('id', logEntry.id);
    }
    throw err;
  }
}

/** TMS에서 직접 관리하는 상태 — 아임웹 동기화 시 덮어쓰지 않음 */
const TMS_MANAGED_STATUSES: OrderStatus[] = ['cancel_pending', 'shipping'];

/** 단건 upsert — 주문 + 품목 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function upsertOrder(supabase: any, imwebOrder: ImwebOrder, prodOrders: ImwebProdOrder[]) {
  const addr = imwebOrder.delivery?.address;
  const pay = imwebOrder.payment;

  // prod-orders에서 상태 가져오기 (첫 번째 품목 기준)
  const imwebStatus = prodOrders.length > 0
    ? mapImwebStatus(prodOrders[0].status)
    : (pay?.payment_time > 0 ? 'pay_done' : 'pay_wait');

  // 기존 주문이 TMS 관리 상태인지 확인 → 있으면 상태/배송정보 보존
  const { data: existing } = await supabase
    .from('orders')
    .select('id, status, invoice_number, courier_code, courier_name, shipped_at, review_requested_at')
    .eq('imweb_order_no', imwebOrder.order_no)
    .single();

  const isTmsManaged = existing && TMS_MANAGED_STATUSES.includes(existing.status);

  const orderData = {
    imweb_order_no: imwebOrder.order_no,
    imweb_order_id: imwebOrder.order_code,
    orderer_name: imwebOrder.orderer?.name || '',
    orderer_phone: imwebOrder.orderer?.call || null,
    orderer_email: imwebOrder.orderer?.email || null,
    recipient_name: addr?.name || '',
    recipient_phone: addr?.phone || null,
    recipient_postcode: addr?.postcode || null,
    recipient_address: addr?.address || null,
    recipient_address_detail: addr?.address_detail || null,
    recipient_memo: imwebOrder.delivery?.memo || null,
    total_price: pay?.total_price || 0,
    delivery_fee: pay?.deliv_price || 0,
    discount_amount: pay?.coupon || 0,
    paid_amount: pay?.payment_amount || 0,
    payment_method: pay?.pay_type || null,
    paid_at: pay?.payment_time
      ? new Date(pay.payment_time * 1000).toISOString()
      : null,
    // TMS 관리 중이면 배송정보/상태 보존
    courier_code: isTmsManaged ? existing.courier_code : (imwebOrder.parcel_code || null),
    invoice_number: isTmsManaged ? existing.invoice_number : (imwebOrder.invoice_no || null),
    status: isTmsManaged ? existing.status : imwebStatus,
    ...(isTmsManaged ? { shipped_at: existing.shipped_at } : {}),
    imweb_raw: imwebOrder as unknown as Record<string, unknown>,
    synced_at: new Date().toISOString(),
    ordered_at: new Date(imwebOrder.order_time * 1000).toISOString(),
  };

  const { data: order, error } = await supabase
    .from('orders')
    .upsert(orderData, { onConflict: 'imweb_order_no' })
    .select('id')
    .single();

  if (error) throw error;

  // 품목 동기화
  if (order && prodOrders.length > 0) {
    // 기존 품목 삭제 후 재삽입
    await supabase.from('order_items').delete().eq('order_id', order.id);

    const items = prodOrders.flatMap((po) =>
      po.items.map((item) => ({
        order_id: order.id,
        imweb_product_no: String(item.prod_no),
        product_name: item.prod_name,
        option_text: item.prod_sku_no || null,
        quantity: item.payment.count,
        unit_price: item.payment.price,
        total_price: item.payment.price * item.payment.count,
      }))
    );

    if (items.length > 0) {
      await supabase.from('order_items').insert(items);
    }
  }

  // 배송완료 전환 감지 → 리뷰 요청 자동 발송
  const wasNotDelivered = !existing || !['delivered', 'confirmed'].includes(existing.status);
  const isNowDelivered = imwebStatus === 'delivered';
  const notYetRequested = !existing?.review_requested_at;

  if (wasNotDelivered && isNowDelivered && notYetRequested && order) {
    const { data: orderItems } = await supabase
      .from('order_items')
      .select('product_name, quantity')
      .eq('order_id', order.id);

    const productNames = (orderItems || [])
      .map((it: { product_name: string; quantity: number }) =>
        it.quantity > 1 ? `${it.product_name} x${it.quantity}` : it.product_name
      )
      .join(', ');

    const phone = imwebOrder.orderer?.call;
    const name = imwebOrder.orderer?.name;
    if (phone && name) {
      await sendNotification({
        template: 'purchase_review_request',
        phone,
        name,
        data: {
          order_uid: order.id,
          product_names: productNames,
          review_type: 'purchase',
          name,
        },
      }).catch((err: unknown) => console.error('[sync] 리뷰 요청 발송 실패:', err));
    }

    await supabase
      .from('orders')
      .update({ review_requested_at: new Date().toISOString() })
      .eq('id', order.id);
  }
}
