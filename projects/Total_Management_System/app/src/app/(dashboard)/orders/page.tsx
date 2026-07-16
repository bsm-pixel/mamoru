'use client';

import { useState, memo } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useOrders, useOrderSync, useOrderCounts } from '@/hooks/use-orders';
import { useIsLg } from '@/hooks/use-grid-mode';
import { formatKRW, formatDateTime, ORDER_STATUS_LABEL, ORDER_STATUS_COLOR } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { DataGrid, type GridColumn } from '@/components/ui/data-grid';
import { OrderDetailPanel } from '@/components/orders/order-detail-panel';
import { RefreshCw, Truck, ShoppingBag } from 'lucide-react';
import { useEscapeKey } from '@/hooks/use-media-query';
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
  const [status, setStatus] = useState('all');
  const [search, setSearch] = useState('');
  const [dateRange, setDateRange] = useState<'all' | 'today' | 'week' | 'month'>('all');
  const [page, setPage] = useState(1);
  const [invoiceOrder, setInvoiceOrder] = useState<Order | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const sync = useOrderSync();
  const { data: counts } = useOrderCounts();

  const isLg = useIsLg();

  useEscapeKey(() => setSelectedId(null), !!selectedId);
  const { data, isLoading } = useOrders({ status, search, dateRange, page, limit: 20 });
  const orders = data?.orders || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  // PC 그리드 컬럼 — 주문일·주문번호·주문자·상태·결제·금액·송장(액션)
  const orderColumns: GridColumn<Order>[] = [
    { key: 'ordered', label: '주문일', render: (o) => <span className="text-neutral-500 whitespace-nowrap tabular-nums text-xs">{formatDateTime(o.ordered_at)}</span> },
    { key: 'no', label: '주문번호', render: (o) => <span className="font-mono text-[11px] text-neutral-500">{o.imweb_order_no}</span> },
    { key: 'name', label: '주문자', render: (o) => <span className="font-semibold text-indigo-black truncate">{o.orderer_name}</span> },
    { key: 'status', label: '상태', render: (o) => <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${ORDER_STATUS_COLOR[o.status] || 'bg-stone-100 text-stone-500'}`}>{ORDER_STATUS_LABEL[o.status] || o.status}</span> },
    { key: 'pay', label: '결제', render: (o) => o.paid_at
      ? <span className="text-emerald-600 text-xs font-medium">결제완료</span>
      : o.paid_amount > 0 ? <span className="text-rose-600 text-xs font-medium">미납</span> : <span className="text-neutral-300">—</span> },
    { key: 'amount', label: '금액', align: 'right', render: (o) => <span className="font-bold tabular-nums text-indigo-black">{formatKRW(o.paid_amount)}</span> },
    { key: 'invoice', label: '송장', align: 'right', render: (o) => o.invoice_number
      ? <span className="text-[11px] text-neutral-600 font-medium whitespace-nowrap">{o.invoice_number}</span>
      : o.status === 'pay_done'
        ? <button onClick={(e) => { e.stopPropagation(); setInvoiceOrder(o); }} className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-stone-900 text-white text-[11px] font-semibold hover:bg-stone-800 transition"><Truck size={11} />송장생성</button>
        : <span className="text-neutral-300">—</span> },
  ];

  // 목록 콘텐츠 — 카드/그리드 공용 (PC·모바일 재사용)
  const listContent = isLoading ? (
    <div className="p-4 space-y-3">
      {Array.from({ length: 5 }).map((_, i) => <Skeleton key={i} className="h-16 w-full" />)}
    </div>
  ) : orders.length === 0 ? (
    <EmptyState icon={ShoppingBag} message="주문이 없습니다" />
  ) : (
    <div className="divide-y divide-neutral-100">
      {orders.map((order) => (
        <OrderRow key={order.id} order={order} isSelected={selectedId === order.id}
          onClick={() => setSelectedId(order.id)} onInvoice={() => setInvoiceOrder(order)} />
      ))}
    </div>
  );

  return (
    <>
      <Topbar title="주문관리" />

      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-3">
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
          <SearchInput
            value={search}
            onChange={(v) => { setSearch(v); setPage(1); }}
            placeholder="주문번호, 이름, 송장번호 검색"
          />
          <select
            value={dateRange}
            onChange={(e) => { setDateRange(e.target.value as typeof dateRange); setPage(1); }}
            className="shrink-0 h-9 px-3 rounded-xl border border-stone-200 bg-white text-xs font-medium text-stone-700 focus:outline-none focus:border-stone-400 transition"
          >
            <option value="all">전체 기간</option>
            <option value="today">오늘</option>
            <option value="week">이번주</option>
            <option value="month">이번달</option>
          </select>
        </div>

        {/* 상태 탭 */}
        <div className="flex gap-1 overflow-x-auto pb-1">
          {STATUS_TABS.map((tab) => (
            <button
              key={tab.value}
              onClick={() => { setStatus(tab.value); setPage(1); }}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
                status === tab.value
                  ? 'bg-stone-900 text-white'
                  : 'bg-stone-100 text-stone-600 hover:bg-stone-200'
              }`}
            >
              {tab.label}
              {counts && counts[tab.value] > 0 && (
                <span className={`ml-1 px-1.5 py-0.5 rounded-full text-[10px] font-bold ${
                  status === tab.value
                    ? 'bg-white/20 text-white'
                    : 'bg-stone-200 text-stone-600'
                }`}>
                  {counts[tab.value]}
                </span>
              )}
            </button>
          ))}
        </div>

        {/* PC: 밀집 그리드 + 우측 상세 (카드보기·토글 폐지 — 항상 그리드) */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-240px)]">
            {/* 좌측: 밀집 그리드 목록 */}
            <div className="flex-1 min-w-0 overflow-auto">
              <Card padding={false}>
                {isLoading ? (
                  <div className="p-4 space-y-3">{Array.from({ length: 5 }).map((_, i) => <Skeleton key={i} className="h-16 w-full" />)}</div>
                ) : (
                  <DataGrid
                    columns={orderColumns}
                    rows={orders}
                    getRowKey={(o) => o.id}
                    selectedKey={selectedId ?? undefined}
                    onSelect={(o) => setSelectedId(o.id)}
                    rowClassName={(o) => (o.status === 'cancelled' ? 'opacity-50' : '')}
                    emptyMessage="주문이 없습니다"
                  />
                )}
              </Card>
              <div className="mt-2">
                <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
              </div>
            </div>

            {/* 우측: 주문 상세 모니터 (고정폭) */}
            <div className="w-[440px] shrink-0 overflow-y-auto">
              {selectedId ? (
                <OrderDetailPanel orderId={selectedId} />
              ) : (
                <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                  <ShoppingBag size={28} className="mb-2 opacity-40" />
                  <p className="text-xs">목록에서 주문을 선택하세요</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* 모바일: 목록만 */}
        {!isLg && (
          <>
            <Card padding={false}>{listContent}</Card>
            <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
          </>
        )}
      </div>

      {/* 모바일 전용 슬라이드 패널 */}
      {!isLg && (
        <SlidePanel
          open={!!selectedId}
          onClose={() => setSelectedId(null)}
          title="주문 상세"
          className="sm:w-[480px]"
        >
          {selectedId && <OrderDetailPanel orderId={selectedId} />}
        </SlidePanel>
      )}

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

const OrderRow = memo(function OrderRow({ order, isSelected, onClick, onInvoice }: { order: Order; isSelected: boolean; onClick: () => void; onInvoice: () => void }) {
  const statusColor = ORDER_STATUS_COLOR[order.status] || 'bg-stone-100 text-stone-500';

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-stone-50 transition ${isSelected ? 'bg-stone-50 border-l-2 border-l-stone-900' : ''}`}
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-stone-800 truncate">
            {order.orderer_name}
          </span>
          <Badge className={statusColor}>
            {ORDER_STATUS_LABEL[order.status] || order.status}
          </Badge>
          {order.paid_at ? (
            <span className="px-1.5 py-0.5 rounded-full text-[10px] font-semibold bg-emerald-50 text-emerald-700">
              결제완료
            </span>
          ) : order.paid_amount > 0 && (
            <span className="px-1.5 py-0.5 rounded-full text-[10px] font-semibold bg-rose-50 text-rose-700">
              미납
            </span>
          )}
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-stone-500">
          <span>{order.imweb_order_no}</span>
          <span>{formatDateTime(order.ordered_at)}</span>
          {order.invoice_number && (
            <span className="text-stone-700 font-medium">{order.invoice_number}</span>
          )}
        </div>
        {order.recipient_memo && (
          <p className="mt-1 text-xs text-stone-400 truncate max-w-[300px]" title={order.recipient_memo}>
            📝 {order.recipient_memo}
          </p>
        )}
      </div>
      <div className="text-right shrink-0 flex flex-col items-end gap-1">
        <span className="text-sm font-bold text-stone-900">{formatKRW(order.paid_amount)}</span>
        {order.status === 'pay_done' && !order.invoice_number && (
          <button
            onClick={(e) => { e.stopPropagation(); onInvoice(); }}
            className="flex items-center gap-1 px-2 py-1 rounded-md bg-stone-900 text-white text-[11px] font-semibold hover:bg-stone-800 transition"
          >
            <Truck size={12} />
            송장생성
          </button>
        )}
      </div>
    </div>
  );
});
