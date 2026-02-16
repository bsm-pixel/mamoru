/**
 * 아임웹 → Supabase 주문 동기화 로직
 * 증분 동기화: 마지막 synced_at 이후 변경분만 가져옴
 */

import { getOrders } from './client';
import type { ImwebOrder } from './types';
import type { OrderStatus } from '@/lib/supabase/types';
import { createServiceClient } from '@/lib/supabase/server';
import { format, subDays } from 'date-fns';

/** 아임웹 status 코드 → TMS status 매핑 */
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

  const fromDate = lastSync?.completed_at
    ? format(new Date(lastSync.completed_at), 'yyyy-MM-dd')
    : format(subDays(new Date(), 30), 'yyyy-MM-dd');  // 첫 동기화: 30일 전부터

  const toDate = format(new Date(), 'yyyy-MM-dd');

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
        order_date_from: fromDate,
        order_date_to: toDate,
        page,
        limit: 50,
      });

      const orders = res.data.list || [];
      if (orders.length === 0) {
        hasMore = false;
        break;
      }

      for (const order of orders) {
        try {
          await upsertOrder(supabase, order);
          totalSynced++;
        } catch (err) {
          errors.push(`주문 ${order.order_no}: ${err}`);
        }
      }

      if (orders.length < 50) {
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

/** 단건 upsert */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function upsertOrder(supabase: any, imwebOrder: ImwebOrder) {
  const orderData = {
    imweb_order_no: imwebOrder.order_no,
    imweb_order_id: imwebOrder.order_code,
    orderer_name: imwebOrder.orderer?.name || '',
    orderer_phone: imwebOrder.orderer?.phone || null,
    orderer_email: imwebOrder.orderer?.email || null,
    recipient_name: imwebOrder.delivery?.name || '',
    recipient_phone: imwebOrder.delivery?.phone || null,
    recipient_postcode: imwebOrder.delivery?.zipcode || null,
    recipient_address: imwebOrder.delivery?.addr || null,
    recipient_address_detail: imwebOrder.delivery?.addr_detail || null,
    recipient_memo: imwebOrder.delivery?.memo || null,
    total_price: imwebOrder.price?.total || 0,
    delivery_fee: imwebOrder.price?.deliv || 0,
    discount_amount: imwebOrder.price?.discount || 0,
    paid_amount: imwebOrder.price?.pay_price || 0,
    payment_method: imwebOrder.pay_type || null,
    paid_at: imwebOrder.pay_time
      ? new Date(imwebOrder.pay_time * 1000).toISOString()
      : null,
    courier_code: imwebOrder.parcel_code || null,
    invoice_number: imwebOrder.invoice_no || null,
    status: mapImwebStatus(imwebOrder.status),
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

  // 주문 품목 동기화
  if (order && imwebOrder.items?.length) {
    // 기존 품목 삭제 후 재삽입
    await supabase.from('order_items').delete().eq('order_id', order.id);

    const items = imwebOrder.items.map((item) => ({
      order_id: order.id,
      imweb_product_no: item.prod_no,
      product_name: item.prod_name,
      option_text: item.options || null,
      quantity: item.qty,
      unit_price: item.price,
      total_price: item.total,
    }));

    await supabase.from('order_items').insert(items);
  }
}
