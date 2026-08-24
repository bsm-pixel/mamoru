'use client';

import { useState } from 'react';
import { useOrder } from '@/hooks/use-orders';
import { OrderSerialModal } from './order-serial-modal';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { DeliveryTracker } from './delivery-tracker';
import { OrderActionBar } from './order-action-bar';
import { formatKRW, formatDateTime, ORDER_STATUS_LABEL, ORDER_STATUS_COLOR } from '@/lib/utils/format';
import { Package, Hash } from 'lucide-react';
import Link from 'next/link';

interface Props {
  orderId: string;
}

export function OrderDetailPanel({ orderId }: Props) {
  const { data, isLoading } = useOrder(orderId);
  const [showSerials, setShowSerials] = useState(false);

  if (isLoading) {
    return <div className="space-y-3"><Skeleton className="h-20" /><Skeleton className="h-32" /><Skeleton className="h-20" /></div>;
  }

  if (!data?.order) {
    return <p className="text-sm text-neutral-400 text-center py-8">주문 정보를 찾을 수 없습니다</p>;
  }

  const { order: o, items, serials } = data;
  const statusColor = ORDER_STATUS_COLOR[o.status] || 'bg-neutral-100 text-neutral-500';
  const addr = [o.recipient_address, o.recipient_address_detail].filter(Boolean).join(' ');
  // 주문자 ≠ 받는분일 때만 주문자 보조표기 (같으면 중복 제거)
  const ordererDiffers = !!o.orderer_name &&
    (o.orderer_name !== o.recipient_name || (o.orderer_phone || '') !== (o.recipient_phone || ''));
  const fee = o.delivery_fee ?? 0;
  const discount = o.discount_amount ?? 0;

  return (
    <div className="space-y-3">
      {/* 헤더 — 주문번호·상태·날짜 */}
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className="flex items-center gap-1.5">
            <h3 className="text-sm font-bold font-mono text-indigo-black truncate">{o.imweb_order_no}</h3>
            {o.is_pickup && (
              <span className="px-1.5 py-0.5 rounded text-[10px] font-semibold bg-violet-50 text-violet-700 shrink-0">직접수령</span>
            )}
          </div>
          <p className="text-[11px] text-neutral-400 mt-0.5">{formatDateTime(o.ordered_at)}</p>
        </div>
        <Badge className={`${statusColor} shrink-0`}>{ORDER_STATUS_LABEL[o.status] || o.status}</Badge>
      </div>

      {/* 받는분 (배송지) — 주문자와 같으면 통합, 다르면 주문자 보조표기 */}
      <div className="rounded-lg bg-neutral-50 p-3 space-y-1.5">
        <div className="flex items-center justify-between gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <span className="text-sm font-semibold text-indigo-black truncate">{o.recipient_name || o.orderer_name}</span>
            {o.recipient_phone && (
              <a href={`tel:${o.recipient_phone}`} className="text-xs text-blue-600 shrink-0">{o.recipient_phone}</a>
            )}
          </div>
          {ordererDiffers && (
            <span className="text-[10px] text-neutral-400 shrink-0 whitespace-nowrap">주문 {o.orderer_name}</span>
          )}
        </div>
        {addr && <p className="text-[13px] text-neutral-600 leading-snug">{addr}</p>}
        {o.recipient_memo && (
          <p className="text-[11px] text-amber-700 bg-amber-100/50 rounded px-2 py-1">📝 {o.recipient_memo}</p>
        )}
      </div>

      {/* 주문 품목 */}
      {items.length > 0 && (
        <div>
          <p className="text-[11px] font-semibold text-neutral-400 mb-1.5 flex items-center gap-1">
            <Package size={11} /> 주문 품목
          </p>
          <div className="space-y-1">
            {items.map((item) => (
              <div key={item.id} className="flex items-center justify-between gap-2 text-sm">
                <div className="min-w-0">
                  <p className="font-medium text-indigo-black truncate">{item.product_name}</p>
                  {item.option_text && <p className="text-[11px] text-neutral-400 truncate">{item.option_text}</p>}
                </div>
                <div className="text-right shrink-0 tabular-nums whitespace-nowrap">
                  <span className="text-[11px] text-neutral-400 mr-1.5">{item.quantity}개</span>
                  <span className="font-semibold text-indigo-black">{formatKRW(item.total_price)}</span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 배정 시리얼 */}
      {items.some((i) => i.product_id) && o.status !== 'cancelled' && (
        <div>
          <div className="flex items-center justify-between mb-1.5">
            <p className="text-[11px] font-semibold text-neutral-400 flex items-center gap-1">
              <Hash size={11} /> 배정 시리얼{serials.length > 0 && <span className="text-neutral-300 ml-0.5">({serials.length})</span>}
            </p>
            <button
              onClick={() => setShowSerials(true)}
              className="text-[11px] text-blue-600 hover:text-blue-700 font-medium"
            >
              {serials.length > 0 ? '수정' : '배정'}
            </button>
          </div>
          {serials.length > 0 ? (
            <div className="flex flex-wrap gap-1">
              {serials.map((s) => (
                <span key={s.id} className="font-mono text-[11px] bg-neutral-100 text-neutral-600 rounded px-1.5 py-0.5">
                  {s.serial_number}
                </span>
              ))}
            </div>
          ) : (
            <p className="text-[11px] text-neutral-400">배정된 시리얼 없음</p>
          )}
        </div>
      )}

      {/* 결제 — 결제금액 강조 + 내역 한 줄 */}
      <div className="rounded-lg bg-neutral-50 p-3">
        <div className="flex items-baseline justify-between">
          <span className="text-xs font-semibold text-neutral-500">결제금액</span>
          <span className="text-lg font-bold text-terracotta tabular-nums">{formatKRW(o.paid_amount)}</span>
        </div>
        {(o.total_price !== o.paid_amount || fee > 0 || discount > 0) && (
          <p className="mt-0.5 text-[11px] text-neutral-400 text-right tabular-nums">
            상품 {formatKRW(o.total_price)}
            {fee > 0 && <> · 배송 {formatKRW(fee)}</>}
            {discount > 0 && <> · 할인 -{formatKRW(discount)}</>}
          </p>
        )}
      </div>

      {/* 배송 추적 — 송장번호 인라인 */}
      {o.invoice_number && o.status !== 'cancelled' && (
        <div>
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-[11px] font-semibold text-neutral-400">배송 추적</span>
            <span className="font-mono text-[11px] text-terracotta">{o.invoice_number}</span>
          </div>
          <DeliveryTracker invNo={o.invoice_number} />
        </div>
      )}

      {/* 상태별 액션 */}
      <OrderActionBar order={o} items={items} />

      {/* 하단 링크 — 한 줄 */}
      <div className="flex items-center justify-between pt-1 text-[11px]">
        <a
          href="https://mamoruscissors63682.imweb.me/admin/shopping/order-v1"
          target="_blank"
          rel="noopener noreferrer"
          className="text-blue-500 hover:text-blue-700"
        >
          아임웹에서 열기 ↗
        </a>
        <Link href={`/orders/${o.id}`} className="text-neutral-400 hover:text-neutral-600">
          상세 페이지 →
        </Link>
      </div>

      {showSerials && (
        <OrderSerialModal orderId={o.id} items={items} serials={serials} onClose={() => setShowSerials(false)} />
      )}
    </div>
  );
}
