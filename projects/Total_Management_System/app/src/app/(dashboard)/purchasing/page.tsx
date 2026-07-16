'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { usePurchaseOrders } from '@/hooks/use-purchasing';
import { useGridMode } from '@/hooks/use-grid-mode';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { DataGrid, GridToggleButton, type GridColumn } from '@/components/ui/data-grid';
import { PurchaseDetailPanel } from '@/components/purchasing/purchase-detail-panel';
import { Plus, Truck } from 'lucide-react';
import type { PurchaseOrder } from '@/lib/supabase/types';

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', ordered: '발주완료', deposit_paid: '선납완료',
  partial: '부분입고', received: '입고완료', balance_paid: '잔금완료', cancelled: '취소',
};
const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600', ordered: 'bg-blue-100 text-blue-700',
  deposit_paid: 'bg-yellow-100 text-yellow-700', partial: 'bg-orange-100 text-orange-700',
  received: 'bg-green-100 text-green-700',
  balance_paid: 'bg-emerald-100 text-emerald-700', cancelled: 'bg-red-100 text-red-700',
};
const STATUS_TABS = [
  { value: '', label: '전체' },
  { value: 'draft', label: '작성중' },
  { value: 'ordered', label: '발주완료' },
  { value: 'deposit_paid', label: '선납완료' },
  { value: 'partial', label: '부분입고' },
  { value: 'received', label: '입고완료' },
  { value: 'balance_paid', label: '잔금완료' },
];

/** 상태 표기 — 입고 전이지만 잔금까지 다 냈으면 '결제완료'(선납완료 오해 방지). PORow·그리드 공용(드리프트 0) */
function poStatusView(po: PurchaseOrder): { label: string; color: string } {
  const paidEarly = !!po.balance_paid_at && (po.status === 'ordered' || po.status === 'deposit_paid');
  return paidEarly
    ? { label: '결제완료', color: 'bg-emerald-100 text-emerald-700' }
    : { label: STATUS_LABEL[po.status] || po.status, color: STATUS_COLOR[po.status] || STATUS_COLOR.draft };
}

/** PC 그리드 컬럼: 발주일·공급처·상태·입고예정·금액 */
const PO_COLUMNS: GridColumn<PurchaseOrder>[] = [
  { key: 'date', label: '발주일', render: (po) => <span className="text-neutral-600 tabular-nums">{formatDate(po.order_date)}</span> },
  { key: 'supplier', label: '공급처', render: (po) => <span className="font-semibold text-indigo-black">{po.supplier_name}</span> },
  { key: 'status', label: '상태', render: (po) => { const v = poStatusView(po); return <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${v.color}`}>{v.label}</span>; } },
  { key: 'expected', label: '입고예정', render: (po) => <span className="text-neutral-500 tabular-nums">{po.expected_date ? formatDate(po.expected_date) : '—'}</span> },
  { key: 'amount', label: '금액', align: 'right', render: (po) => <span className="font-bold tabular-nums text-indigo-black">{formatKRW(po.total_amount)}</span> },
];

export default function PurchasingPage() {
  const router = useRouter();
  const [statusFilter, setStatusFilter] = useState('');
  const [search, setSearch] = useState('');
  const [dateRange, setDateRange] = useState<'all' | 'today' | 'week' | 'month'>('all');
  const [page, setPage] = useState(1);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const limit = 20;

  const { isLg, gridMode, toggleGrid } = useGridMode('purchasing-pc-grid');

  const { data, isLoading } = usePurchaseOrders({
    status: statusFilter || undefined,
    search: search || undefined,
    dateRange,
    page,
    limit,
  });
  const orders = data?.orders || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);

  const listContent = isLoading ? (
    <div className="p-4 space-y-3">
      {Array.from({ length: 5 }).map((_, i) => (
        <Skeleton key={i} className="h-16 w-full" />
      ))}
    </div>
  ) : orders.length === 0 ? (
    <EmptyState icon={Truck} message="발주 내역이 없습니다" />
  ) : (
    <div className="divide-y divide-neutral-100">
      {orders.map((po) => (
        <PORow key={po.id} po={po} isSelected={selectedId === po.id} onClick={() => setSelectedId(po.id)} />
      ))}
    </div>
  );

  return (
    <>
      <Topbar title="매입관리" />

      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-3">
        <div className="flex items-center gap-3">
          <Button size="sm" onClick={() => router.push('/purchasing/new')}>
            <Plus size={14} />
            발주 작성
          </Button>
          <SearchInput
            value={search}
            onChange={(v) => { setSearch(v); setPage(1); }}
            placeholder="발주번호, 매입처 검색"
          />
          <select
            value={dateRange}
            onChange={(e) => { setDateRange(e.target.value as typeof dateRange); setPage(1); }}
            className="shrink-0 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-xs font-medium text-neutral-600 focus:outline-none focus:ring-2 focus:ring-stone-400"
          >
            <option value="all">전체 기간</option>
            <option value="today">오늘</option>
            <option value="week">이번주</option>
            <option value="month">이번달</option>
          </select>
          <div className="ml-auto">
            <GridToggleButton isLg={isLg} gridMode={gridMode} onToggle={toggleGrid} />
          </div>
        </div>

        <div className="flex gap-1.5 overflow-x-auto pb-1">
          {STATUS_TABS.map((tab) => (
            <button
              key={tab.value}
              onClick={() => { setStatusFilter(tab.value); setPage(1); }}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                statusFilter === tab.value
                  ? 'bg-stone-900 text-white'
                  : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* PC: 2열 마스터-디테일 (그리드모드 시 목록 넓게/상세 420px 반전) */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-240px)]">
            <div className={`${gridMode ? 'flex-1 min-w-0' : 'w-[40%] shrink-0'} overflow-y-auto`}>
              <Card padding={false}>
                {gridMode && !isLoading ? (
                  <DataGrid
                    columns={PO_COLUMNS}
                    rows={orders}
                    getRowKey={(po) => po.id}
                    selectedKey={selectedId ?? undefined}
                    onSelect={(po) => setSelectedId(po.id)}
                    rowClassName={(po) => (po.status === 'cancelled' ? 'opacity-50' : '')}
                    emptyMessage="발주 내역이 없습니다"
                  />
                ) : (
                  listContent
                )}
              </Card>
              <div className="mt-2">
                <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
              </div>
            </div>
            <div className={`${gridMode ? 'w-[420px] shrink-0' : 'flex-1 min-w-0'} overflow-y-auto`}>
              {selectedId ? (
                <PurchaseDetailPanel purchaseId={selectedId} />
              ) : (
                <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                  <Truck size={28} className="mb-2 opacity-40" />
                  <p className="text-xs">목록에서 발주를 선택하세요</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* 모바일 */}
        {!isLg && (
          <>
            <Card padding={false}>{listContent}</Card>
            <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
          </>
        )}
      </div>

      {/* 모바일 전용 슬라이드 패널 */}
      {!isLg && (
        <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="발주 상세" className="sm:w-[640px]">
          {selectedId && <PurchaseDetailPanel purchaseId={selectedId} />}
        </SlidePanel>
      )}
    </>
  );
}

function PORow({ po, isSelected, onClick }: { po: PurchaseOrder; isSelected: boolean; onClick: () => void }) {
  return (
    <div onClick={onClick} className={`flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-stone-50/60 transition ${isSelected ? 'bg-stone-100 border-l-2 border-l-stone-900' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-stone-900 truncate">{po.supplier_name}</span>
          <Badge className={poStatusView(po).color}>{poStatusView(po).label}</Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{formatDate(po.order_date)}</span>
          {po.expected_date && <span>예정: {formatDate(po.expected_date)}</span>}
        </div>
      </div>
      <div className="text-right shrink-0">
        {(() => {
          const cur = (po as Record<string, unknown>).currency as string | undefined;
          const rate = (po as Record<string, unknown>).exchange_rate as number | undefined;
          return cur && cur !== 'KRW' ? (
            <p className="text-[10px] text-neutral-400">{cur} 환율 {rate?.toLocaleString()}</p>
          ) : null;
        })()}
        <p className="text-sm font-bold">{formatKRW(po.total_amount)}</p>
        {po.balance_paid_at && po.status !== 'balance_paid' ? (
          <p className="text-xs text-emerald-600">잔금 지불완료 ✓</p>
        ) : po.deposit_amount > 0 && po.status !== 'balance_paid' ? (
          <p className="text-xs text-neutral-500">선납 {formatKRW(po.deposit_amount)}</p>
        ) : null}
      </div>
    </div>
  );
}
