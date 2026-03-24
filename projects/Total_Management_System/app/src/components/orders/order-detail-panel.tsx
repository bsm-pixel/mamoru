'use client';

import { useOrder } from '@/hooks/use-orders';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { DeliveryTracker } from './delivery-tracker';
import { OrderActionBar } from './order-action-bar';
import { formatKRW, formatDateTime, ORDER_STATUS_LABEL, ORDER_STATUS_COLOR } from '@/lib/utils/format';
import { User, Phone, MapPin, Package, CreditCard } from 'lucide-react';
import Link from 'next/link';

interface Props {
  orderId: string;
}

export function OrderDetailPanel({ orderId }: Props) {
  const { data, isLoading } = useOrder(orderId);

  if (isLoading) {
    return <div className="space-y-3"><Skeleton className="h-20" /><Skeleton className="h-32" /><Skeleton className="h-20" /></div>;
  }

  if (!data?.order) {
    return <p className="text-sm text-neutral-400 text-center py-8">주문 정보를 찾을 수 없습니다</p>;
  }

  const { order: o, items } = data;
  const statusColor = ORDER_STATUS_COLOR[o.status] || 'bg-neutral-100 text-neutral-500';

  return (
    <div className="space-y-4">
      {/* 헤더 */}
      <div>
        <div className="flex items-center gap-2 mb-1">
          <h3 className="text-base font-bold">{o.imweb_order_no}</h3>
          <Badge className={statusColor}>{ORDER_STATUS_LABEL[o.status] || o.status}</Badge>
        </div>
        <p className="text-xs text-neutral-500">{formatDateTime(o.ordered_at)}</p>
      </div>

      {/* 주문자 정보 */}
      <div className="bg-neutral-50 rounded-lg p-3 space-y-2">
        <div className="flex items-center gap-2 text-sm">
          <User size={14} className="text-neutral-400" />
          <span className="font-medium">{o.orderer_name}</span>
        </div>
        {o.orderer_phone && (
          <div className="flex items-center gap-2 text-sm">
            <Phone size={14} className="text-neutral-400" />
            <a href={`tel:${o.orderer_phone}`} className="text-blue-600">{o.orderer_phone}</a>
          </div>
        )}
      </div>

      {/* 배송지 */}
      <div className="bg-neutral-50 rounded-lg p-3 space-y-2">
        <p className="text-xs font-semibold text-neutral-500 mb-1">배송지</p>
        <div className="flex items-center gap-2 text-sm">
          <User size={14} className="text-neutral-400" />
          <span>{o.recipient_name}</span>
          {o.recipient_phone && (
            <a href={`tel:${o.recipient_phone}`} className="text-blue-600 text-xs">{o.recipient_phone}</a>
          )}
        </div>
        {o.recipient_address && (
          <div className="flex items-start gap-2 text-sm">
            <MapPin size={14} className="text-neutral-400 shrink-0 mt-0.5" />
            <span className="text-neutral-700">
              {o.recipient_address} {o.recipient_address_detail || ''}
            </span>
          </div>
        )}
        {o.recipient_memo && (
          <p className="text-xs text-neutral-500 ml-6">📝 {o.recipient_memo}</p>
        )}
      </div>

      {/* 주문 품목 */}
      {items.length > 0 && (
        <div>
          <p className="text-xs font-semibold text-neutral-500 mb-2 flex items-center gap-1">
            <Package size={12} /> 주문 품목
          </p>
          <div className="space-y-1.5">
            {items.map((item) => (
              <div key={item.id} className="flex items-center justify-between text-sm">
                <div className="min-w-0">
                  <p className="font-medium truncate">{item.product_name}</p>
                  {item.option_text && <p className="text-xs text-neutral-500">{item.option_text}</p>}
                </div>
                <div className="text-right shrink-0 ml-2">
                  <p className="text-xs text-neutral-500">{item.quantity}개</p>
                  <p className="font-semibold">{formatKRW(item.total_price)}</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* 결제 정보 */}
      <div className="bg-neutral-50 rounded-lg p-3">
        <p className="text-xs font-semibold text-neutral-500 mb-2 flex items-center gap-1">
          <CreditCard size={12} /> 결제 정보
        </p>
        <div className="space-y-1 text-sm">
          <div className="flex justify-between">
            <span className="text-neutral-500">상품금액</span>
            <span>{formatKRW(o.total_price)}</span>
          </div>
          {(o.delivery_fee ?? 0) > 0 && (
            <div className="flex justify-between">
              <span className="text-neutral-500">배송비</span>
              <span>{formatKRW(o.delivery_fee)}</span>
            </div>
          )}
          {(o.discount_amount ?? 0) > 0 && (
            <div className="flex justify-between">
              <span className="text-neutral-500">할인</span>
              <span className="text-red-500">-{formatKRW(o.discount_amount)}</span>
            </div>
          )}
          <div className="flex justify-between pt-1 border-t border-neutral-200 font-bold">
            <span>결제금액</span>
            <span className="text-terracotta">{formatKRW(o.paid_amount)}</span>
          </div>
        </div>
      </div>

      {/* 배송 추적 */}
      {o.invoice_number && !['cancelled'].includes(o.status) && (
        <div>
          <p className="text-xs font-semibold text-neutral-500 mb-2">배송 추적</p>
          <p className="text-xs text-neutral-500 mb-2">송장번호: <span className="font-mono text-terracotta">{o.invoice_number}</span></p>
          <DeliveryTracker invNo={o.invoice_number} />
        </div>
      )}

      {/* 상태별 액션 */}
      <OrderActionBar order={o} items={items} />

      {/* 상세 페이지 링크 */}
      <Link
        href={`/orders/${o.id}`}
        className="block text-center text-xs text-neutral-400 hover:text-neutral-600 py-2"
      >
        상세 페이지에서 보기 →
      </Link>
    </div>
  );
}
