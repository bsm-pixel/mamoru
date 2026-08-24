/**
 * 아임웹 → Supabase 주문 동기화 로직
 * 증분 동기화: 마지막 synced_at 이후 변경분만 가져옴
 */

import { getOrders, getOrder, getProdOrders, updateImwebStock } from './client';
import type { ImwebOrder, ImwebProdOrder } from './types';
import type { OrderStatus } from '@/lib/supabase/types';
import { createServiceClient } from '@/lib/supabase/server';
import { matchOrCreateCustomer } from '@/lib/customer/match-or-create';
import { sendNotification } from '@/lib/notification/make-webhook';
import { subDays } from 'date-fns';
import { after } from 'next/server';

/** 재고 차감이 필요한 상태 (결제 완료 이상) */
const STOCK_DEDUCT_STATUSES: OrderStatus[] = ['pay_done', 'preparing', 'ready_to_ship', 'shipping', 'delivered', 'confirmed'];
/** 재고 복구가 필요한 상태 (취소/환불) */
const STOCK_RESTORE_STATUSES: OrderStatus[] = ['cancelled', 'refunded'];

/** 아임웹 prod-order status → TMS status 매핑 */
function mapImwebStatus(status: string): OrderStatus {
  const map: Record<string, OrderStatus> = {
    BEFORE_DEPOSIT: 'pay_wait',
    PAY_COMPLETE: 'pay_done',
    PREPARE: 'preparing',
    STANDBY: 'ready_to_ship',      // 128: 아임웹 배송대기 → TMS 배송대기
    DELIVERY_READY: 'ready_to_ship',
    DELIVERY: 'shipping',
    DELIVERING: 'shipping',
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

/**
 * 웹훅용 단건 동기화 — 주문번호 1건만 아임웹에서 재조회하여 TMS에 반영.
 * 기존 upsertOrder 재사용 → imweb_order_no 유니크로 크론과 겹쳐도 멱등(중복 없음).
 * 웹훅 페이로드를 신뢰하지 않고 아임웹 API로 재조회하므로 위조 페이로드 방어도 겸함.
 */
export async function syncSingleOrder(
  orderNo: string
): Promise<{ success: boolean; order_no: string; error?: string }> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  try {
    const orderRes = await getOrder(orderNo);
    const order = orderRes.data;
    if (!order) {
      return { success: false, order_no: orderNo, error: 'order not found in imweb' };
    }
    const prodRes = await getProdOrders(orderNo);
    const prodOrders = prodRes.data || [];
    await upsertOrder(supabase, order, prodOrders);
    return { success: true, order_no: orderNo };
  } catch (err) {
    console.error(`[webhook] 단건 동기화 실패 ${orderNo}:`, err);
    return { success: false, order_no: orderNo, error: String(err) };
  }
}

/** TMS에서 직접 관리하는 상태 — 아임웹 동기화 시 덮어쓰지 않음 */
const TMS_MANAGED_STATUSES: OrderStatus[] = ['cancel_pending', 'ready_to_ship', 'shipping'];
/** 배송완료 이후 상태 — shipping 보호를 해제하여 delivered 전환 허용 */
const DELIVERED_OR_LATER: OrderStatus[] = ['delivered', 'confirmed', 'cancelled', 'refund_request', 'refunded'];

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
    .select('id, status, invoice_number, courier_code, courier_name, shipped_at, review_requested_at, stock_deducted')
    .eq('imweb_order_no', imwebOrder.order_no)
    .single();

  // shipping 보호: 아임웹이 delivered 이상으로 진행됐으면 보호 해제 (배송완료 전환 허용)
  const isTmsManaged = existing
    && TMS_MANAGED_STATUSES.includes(existing.status)
    && !DELIVERED_OR_LATER.includes(imwebStatus);

  // 128: 주문을 TMS 고객과 연결 — 전화번호로 matchOrCreate (repairs/consultations 와 동일 패턴)
  //   온라인 전용 고객이면 신규 생성(customer_type='online'), 기존 고객이면 그 id 로 연결.
  const { customerId } = await matchOrCreateCustomer(supabase, {
    phone: imwebOrder.orderer?.call || '',
    name: imwebOrder.orderer?.name || '',
    source: 'imweb',
    extra: { customerType: 'online' },
  });

  const orderData = {
    imweb_order_no: imwebOrder.order_no,
    imweb_order_id: imwebOrder.order_code,
    customer_id: customerId,
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

  // 신규 주문이면 관리자 푸시 (기존 주문 업데이트는 푸시 안 함)
  const isNewOrder = !existing;
  if (isNewOrder && imwebStatus !== 'cancelled') {
    const firstItem = prodOrders[0]?.items?.[0];
    const itemCount = prodOrders.reduce((s, po) => s + po.items.length, 0);
    const body = firstItem
      ? `${imwebOrder.orderer?.name || '고객'}님 — ${firstItem.prod_name}${itemCount > 1 ? ` 외 ${itemCount - 1}건` : ''}`
      : `${imwebOrder.orderer?.name || '고객'}님 주문`;
    // 🔴 완주 보장 — after()로 넘겨야 응답 후에도 푸시가 끝까지 발송된다 (fire-and-forget 누락 방지, 2026-08-01)
    const deliverPush = async () => {
      const { sendPushToAll } = await import('@/lib/firebase/send-push');
      await sendPushToAll({ title: '새 아임웹 주문 📦', body, url: '/orders', tag: `mamoru-order-${imwebOrder.order_no}` });
    };
    try {
      after(deliverPush);
    } catch {
      await deliverPush().catch(() => {});
    }
  }

  // 품목 동기화 — upsert로 트랜잭션 보호 (delete→insert 사이 에러 시 데이터 유실 방지)
  if (order && prodOrders.length > 0) {
    const rawItems = prodOrders.flatMap((po) =>
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
    // imweb_product_no → products.id 매핑을 채운다 (시리얼 배정·재고 매칭이 product_id 기준이므로 필수)
    const prodNos = [...new Set(rawItems.map((i) => i.imweb_product_no))];
    const prodIdMap: Record<string, string> = {};
    if (prodNos.length > 0) {
      const { data: prods } = await supabase
        .from('products')
        .select('id, imweb_product_no')
        .in('imweb_product_no', prodNos);
      (prods || []).forEach((p: { id: string; imweb_product_no: string | null }) => {
        if (p.imweb_product_no) prodIdMap[String(p.imweb_product_no)] = p.id;
      });
    }
    const items = rawItems.map((i) => ({ ...i, product_id: prodIdMap[i.imweb_product_no] || null }));

    if (items.length > 0) {
      // upsert: order_id + imweb_product_no 기준 (중복 시 업데이트)
      await supabase.from('order_items').upsert(items, {
        onConflict: 'order_id,imweb_product_no',
      });

      // upsert 후 이번 동기화에 포함되지 않은 이전 품목 삭제
      const currentProductNos = items.map((i) => i.imweb_product_no);
      await supabase
        .from('order_items')
        .delete()
        .eq('order_id', order.id)
        .not('imweb_product_no', 'in', `(${currentProductNos.map((n) => `"${n}"`).join(',')})`);
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

  // 온라인 주문 재고 차감/복구
  if (order) {
    const finalStatus = isTmsManaged ? existing.status : imwebStatus;
    const alreadyDeducted = existing?.stock_deducted === true;

    // 결제완료 이상 & 아직 차감 안 됨 → 재고 차감
    if (STOCK_DEDUCT_STATUSES.includes(finalStatus) && !alreadyDeducted) {
      await adjustOrderStock(supabase, order.id, 'deduct');
      await supabase.from('orders').update({ stock_deducted: true }).eq('id', order.id);
    }

    // 취소/환불 & 이미 차감됨 → 재고 복구
    if (STOCK_RESTORE_STATUSES.includes(finalStatus) && alreadyDeducted) {
      await adjustOrderStock(supabase, order.id, 'restore');
      await supabase.from('orders').update({ stock_deducted: false }).eq('id', order.id);
    }
  }
}

/** 주문 품목 기준 TMS 재고 차감/복구 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function adjustOrderStock(supabase: any, orderId: string, action: 'deduct' | 'restore') {
  const { data: items } = await supabase
    .from('order_items')
    .select('imweb_product_no, quantity')
    .eq('order_id', orderId);

  if (!items || items.length === 0) return;

  for (const item of items) {
    if (!item.imweb_product_no) continue;

    // imweb_product_no로 TMS 제품 찾기
    const { data: product } = await supabase
      .from('products')
      .select('id, stock_quantity, raw_stock')
      .eq('imweb_product_no', item.imweb_product_no)
      .single();

    if (!product) continue;

    const qty = item.quantity;
    const stockDelta = action === 'deduct' ? -qty : qty;
    let rawDelta = action === 'deduct' ? -qty : qty; // 기본: stock과 동일(loose 가정)

    // 복구(취소/환불) 시: 이 주문·제품에 배정된 시리얼이 있으면 in_stock 복원 + raw 복구량 보정.
    // 배정 API가 raw_stock 을 +시리얼수 해뒀으므로, 복구도 그만큼 덜 되돌려야 불변식(stock=raw+시리얼) 유지.
    //   rawDelta = qty − 배정시리얼수  (시리얼 없으면 = qty, 기존 동작과 동일)
    if (action === 'restore') {
      const { data: soldSerials } = await supabase
        .from('product_serials')
        .select('id, previous_zone, warehouse_zone')
        .eq('order_id', orderId)
        .eq('product_id', product.id)
        .eq('status', 'sold');
      const serials = (soldSerials || []) as Array<{ id: string; previous_zone: string | null; warehouse_zone: string | null }>;
      for (const s of serials) {
        await supabase
          .from('product_serials')
          .update({
            status: 'in_stock',
            warehouse_zone: s.previous_zone || s.warehouse_zone || 'ready',
            sold_via: null, order_id: null, sold_at: null, sold_to_name: null, sold_to_phone: null,
          })
          .eq('id', s.id)
          .eq('order_id', orderId);
      }
      rawDelta = qty - serials.length;
    }

    const newQty = Math.max(0, (product.stock_quantity || 0) + stockDelta);
    const newRaw = Math.max(0, (product.raw_stock || 0) + rawDelta);

    await supabase
      .from('products')
      .update({ stock_quantity: newQty, raw_stock: newRaw, updated_at: new Date().toISOString() })
      .eq('id', product.id);

    // 아임웹 재고 동기화 — stockDelta(±qty, 시리얼과 무관: 아임웹은 우리 시리얼을 모름)
    try {
      await updateImwebStock(Number(item.imweb_product_no), stockDelta);
    } catch (e) {
      console.error(`[sync] 아임웹 재고 동기화 실패: ${item.imweb_product_no}`, e);
    }

    console.log(`[sync] 재고 ${action}: ${item.imweb_product_no} stock=${newQty} raw=${newRaw}`);
  }
}

/** TMS에서 주문 취소 시 재고·시리얼 즉시 복구 (adjustOrderStock restore 재사용 — 시리얼 in_stock 복원 + raw 보정 포함). */
export async function restoreOrderStock(orderId: string): Promise<void> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const supabase: any = createServiceClient();
  await adjustOrderStock(supabase, orderId, 'restore');
}
