'use client';

import { useState } from 'react';
import { useParams, useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { InvoiceModal } from '@/components/orders/invoice-modal';
import { useOrder, useCancelInvoice } from '@/hooks/use-orders';
import {
  formatKRW,
  formatDateTime,
  formatPhone,
  ORDER_STATUS_LABEL,
  ORDER_STATUS_COLOR,
} from '@/lib/utils/format';
import { ArrowLeft, Truck, X, Package } from 'lucide-react';

export default function OrderDetailPage() {
  const params = useParams();
  const router = useRouter();
  const id = params.id as string;
  const { data, isLoading } = useOrder(id);
  const cancelInvoice = useCancelInvoice();
  const [invoiceOpen, setInvoiceOpen] = useState(false);

  if (isLoading) {
    return (
      <>
        <Topbar title="주문 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-8 w-32" />
          <Skeleton className="h-40 w-full" />
          <Skeleton className="h-40 w-full" />
        </div>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <Topbar title="주문 상세" />
        <div className="flex items-center justify-center h-60 text-neutral-400">
          주문을 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { order, items } = data;
  const statusColor = ORDER_STATUS_COLOR[order.status] || 'bg-neutral-100 text-neutral-500';

  return (
    <>
      <Topbar title="주문 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4 max-w-3xl">
        {/* 뒤로가기 */}
        <button
          onClick={() => router.back()}
          className="flex items-center gap-1 text-sm text-neutral-500 hover:text-indigo-black transition"
        >
          <ArrowLeft size={16} />
          목록으로
        </button>

        {/* 주문 헤더 */}
        <div className="flex items-center justify-between">
          <div>
            <h2 className="text-lg font-bold">{order.imweb_order_no}</h2>
            <p className="text-xs text-neutral-500 mt-0.5">
              {formatDateTime(order.ordered_at)}
            </p>
          </div>
          <Badge className={statusColor}>
            {ORDER_STATUS_LABEL[order.status] || order.status}
          </Badge>
        </div>

        {/* 주문자 정보 */}
        <Card>
          <CardHeader>
            <CardTitle>주문자</CardTitle>
          </CardHeader>
          <dl className="grid grid-cols-2 gap-y-2 text-sm">
            <dt className="text-neutral-500">이름</dt>
            <dd className="font-medium">{order.orderer_name}</dd>
            <dt className="text-neutral-500">전화</dt>
            <dd>{formatPhone(order.orderer_phone)}</dd>
            <dt className="text-neutral-500">이메일</dt>
            <dd>{order.orderer_email || '-'}</dd>
          </dl>
        </Card>

        {/* 배송 정보 */}
        <Card>
          <CardHeader>
            <CardTitle>배송지</CardTitle>
          </CardHeader>
          <dl className="grid grid-cols-2 gap-y-2 text-sm">
            <dt className="text-neutral-500">수신자</dt>
            <dd className="font-medium">{order.recipient_name}</dd>
            <dt className="text-neutral-500">전화</dt>
            <dd>{formatPhone(order.recipient_phone)}</dd>
            <dt className="text-neutral-500">우편번호</dt>
            <dd>{order.recipient_postcode || '-'}</dd>
            <dt className="text-neutral-500">주소</dt>
            <dd className="col-span-2">
              {order.recipient_address} {order.recipient_address_detail}
            </dd>
            {order.recipient_memo && (
              <>
                <dt className="text-neutral-500">메모</dt>
                <dd>{order.recipient_memo}</dd>
              </>
            )}
          </dl>
        </Card>

        {/* 주문 품목 */}
        <Card>
          <CardHeader>
            <CardTitle>주문 품목</CardTitle>
          </CardHeader>
          {items.length === 0 ? (
            <p className="text-sm text-neutral-400">품목 정보 없음</p>
          ) : (
            <div className="divide-y divide-neutral-100">
              {items.map((item) => (
                <div key={item.id} className="flex items-center gap-3 py-3">
                  <div className="w-8 h-8 rounded bg-warm-ivory flex items-center justify-center">
                    <Package size={16} className="text-neutral-400" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium truncate">{item.product_name}</p>
                    {item.option_text && (
                      <p className="text-xs text-neutral-500">{item.option_text}</p>
                    )}
                  </div>
                  <div className="text-right text-sm">
                    <span className="text-neutral-500">{item.quantity}개</span>
                    <span className="ml-2 font-semibold">{formatKRW(item.total_price)}</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>

        {/* 결제 정보 */}
        <Card>
          <CardHeader>
            <CardTitle>결제</CardTitle>
          </CardHeader>
          <dl className="space-y-2 text-sm">
            <div className="flex justify-between">
              <dt className="text-neutral-500">상품금액</dt>
              <dd>{formatKRW(order.total_price)}</dd>
            </div>
            <div className="flex justify-between">
              <dt className="text-neutral-500">배송비</dt>
              <dd>{formatKRW(order.delivery_fee)}</dd>
            </div>
            {order.discount_amount > 0 && (
              <div className="flex justify-between">
                <dt className="text-neutral-500">할인</dt>
                <dd className="text-error">-{formatKRW(order.discount_amount)}</dd>
              </div>
            )}
            <div className="border-t border-neutral-200 pt-2 flex justify-between">
              <dt className="font-semibold">결제금액</dt>
              <dd className="font-bold text-base">{formatKRW(order.paid_amount)}</dd>
            </div>
          </dl>
        </Card>

        {/* 배송/송장 */}
        <Card>
          <CardHeader>
            <CardTitle>배송</CardTitle>
          </CardHeader>

          {order.invoice_number ? (
            <div className="space-y-3">
              <div className="flex items-center justify-between p-3 rounded-lg bg-warm-ivory">
                <div>
                  <p className="text-xs text-neutral-500">{order.courier_name || '롯데택배'}</p>
                  <p className="text-sm font-bold mt-0.5">{order.invoice_number}</p>
                </div>
                <Button
                  variant="danger"
                  size="sm"
                  onClick={() =>
                    cancelInvoice.mutate({
                      invNo: order.invoice_number!,
                      orderId: order.id,
                    })
                  }
                  disabled={cancelInvoice.isPending}
                >
                  <X size={14} />
                  {cancelInvoice.isPending ? '취소 중...' : '송장 취소'}
                </Button>
              </div>
              {order.shipped_at && (
                <p className="text-xs text-neutral-500">
                  출고일: {formatDateTime(order.shipped_at)}
                </p>
              )}
            </div>
          ) : (
            <Button onClick={() => setInvoiceOpen(true)} className="w-full">
              <Truck size={16} />
              송장 생성
            </Button>
          )}
        </Card>
      </div>

      {/* 송장 생성 모달 */}
      <InvoiceModal
        open={invoiceOpen}
        onClose={() => setInvoiceOpen(false)}
        order={order}
      />
    </>
  );
}
