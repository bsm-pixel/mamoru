'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useOrders, useOrderSync, useOrderCounts } from '@/hooks/use-orders';
import { formatKRW, formatDateTime, ORDER_STATUS_LABEL, ORDER_STATUS_COLOR } from '@/lib/utils/format';
import { RefreshCw, Search, ChevronLeft, ChevronRight, Truck } from 'lucide-react';
import { InvoiceModal } from '@/components/orders/invoice-modal';
import type { Order } from '@/lib/supabase/types';

const STATUS_TABS = [
  { value: 'all', label: '전체' },
  { value: 'pay_done', label: '결제완료' },
  { value: 'preparing', label: '준비중' },
  { value: 'shipping', label: '배송중' },
  { value: 'delivered', label: '배송완료' },
  { value: 'cancel_pending', label: 'ALPS취소대기' },
  { value: 'cancelled', label: '취소' },
];

export default function OrdersPage() {
  const router = useRouter();
  const [status, setStatus] = useState('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [invoiceOrder, setInvoiceOrder] = useState<Order | null>(null);
  const sync = useOrderSync();
  const { data: counts } = useOrderCounts();

  const { data, isLoading } = useOrders({ status, search, page, limit: 20 });
  const orders = data?.orders || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  return (
    <>
      <Topbar title="주문관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 상단: 동기화 + 검색 */}
        <div className="flex items-center gap-3">
          <Button
            variant="secondary"
            size="sm"
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
          >
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '동기화 중...' : '아임웹 동기화'}
          </Button>

          <div className="flex-1 relative">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => { setSearch(e.target.value); setPage(1); }}
              placeholder="주문번호, 이름, 송장번호 검색"
              className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
            />
          </div>
        </div>

        {/* 상태 탭 */}
        <div className="flex gap-1 overflow-x-auto pb-1">
          {STATUS_TABS.map((tab) => (
            <button
              key={tab.value}
              onClick={() => { setStatus(tab.value); setPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
                status === tab.value
                  ? 'bg-terracotta text-cream'
                  : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
              }`}
            >
              {tab.label}
              {counts && counts[tab.value] > 0 && (
                <span className={`ml-1 px-1.5 py-0.5 rounded-full text-[10px] font-bold ${
                  status === tab.value
                    ? 'bg-cream/20 text-cream'
                    : 'bg-neutral-200 text-neutral-600'
                }`}>
                  {counts[tab.value]}
                </span>
              )}
            </button>
          ))}
        </div>

        {/* 주문 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : orders.length === 0 ? (
            <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
              주문이 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {orders.map((order) => (
                <OrderRow
                  key={order.id}
                  order={order}
                  onClick={() => router.push(`/orders/${order.id}`)}
                  onInvoice={() => setInvoiceOrder(order)}
                />
              ))}
            </div>
          )}
        </Card>

        {/* 페이지네이션 */}
        {totalPages > 1 && (
          <div className="flex items-center justify-center gap-2">
            <Button
              variant="ghost"
              size="sm"
              disabled={page <= 1}
              onClick={() => setPage(page - 1)}
            >
              <ChevronLeft size={16} />
            </Button>
            <span className="text-sm text-neutral-500">
              {page} / {totalPages}
            </span>
            <Button
              variant="ghost"
              size="sm"
              disabled={page >= totalPages}
              onClick={() => setPage(page + 1)}
            >
              <ChevronRight size={16} />
            </Button>
          </div>
        )}
      </div>

      {/* 인라인 송장 생성 모달 */}
      {invoiceOrder && (
        <InvoiceModal
          open={!!invoiceOrder}
          onClose={() => setInvoiceOrder(null)}
          order={invoiceOrder}
        />
      )}
    </>
  );
}

function OrderRow({ order, onClick, onInvoice }: { order: Order; onClick: () => void; onInvoice: () => void }) {
  const statusColor = ORDER_STATUS_COLOR[order.status] || 'bg-neutral-100 text-neutral-500';

  return (
    <div
      onClick={onClick}
      className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">
            {order.orderer_name}
          </span>
          <Badge className={statusColor}>
            {ORDER_STATUS_LABEL[order.status] || order.status}
          </Badge>
          {/* R4: 결제/미납 칩 */}
          {order.paid_at ? (
            <span className="px-1.5 py-0.5 rounded-full text-[10px] font-semibold bg-green-100 text-green-700">
              결제완료
            </span>
          ) : order.paid_amount > 0 && (
            <span className="px-1.5 py-0.5 rounded-full text-[10px] font-semibold bg-red-100 text-red-700">
              미납
            </span>
          )}
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{order.imweb_order_no}</span>
          <span>{formatDateTime(order.ordered_at)}</span>
          {order.invoice_number && (
            <span className="text-terracotta">{order.invoice_number}</span>
          )}
        </div>
        {/* R4: 배송 메모 말줄임 + 호버 전문 */}
        {order.recipient_memo && (
          <p
            className="mt-1 text-xs text-neutral-400 truncate max-w-[300px]"
            title={order.recipient_memo}
          >
            📝 {order.recipient_memo}
          </p>
        )}
      </div>
      <div className="text-right shrink-0 flex flex-col items-end gap-1">
        <span className="text-sm font-bold">{formatKRW(order.paid_amount)}</span>
        {order.status === 'pay_done' && !order.invoice_number && (
          <button
            onClick={(e) => { e.stopPropagation(); onInvoice(); }}
            className="flex items-center gap-1 px-2 py-1 rounded-md bg-terracotta/10 text-terracotta text-[11px] font-semibold hover:bg-terracotta/20 transition"
          >
            <Truck size={12} />
            송장생성
          </button>
        )}
      </div>
    </div>
  );
}
