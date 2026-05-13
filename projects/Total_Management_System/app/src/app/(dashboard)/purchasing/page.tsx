'use client';

import { useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { usePurchaseOrders } from '@/hooks/use-purchasing';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { PurchaseDetailPanel } from '@/components/purchasing/purchase-detail-panel';
import { Plus, Truck } from 'lucide-react';
import type { PurchaseOrder } from '@/lib/supabase/types';

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', ordered: '발주완료', deposit_paid: '선납완료',
  received: '입고완료', balance_paid: '잔금완료', cancelled: '취소',
};
const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600', ordered: 'bg-blue-100 text-blue-700',
  deposit_paid: 'bg-yellow-100 text-yellow-700', received: 'bg-green-100 text-green-700',
  balance_paid: 'bg-emerald-100 text-emerald-700', cancelled: 'bg-red-100 text-red-700',
};
const STATUS_TABS = [
  { value: '', label: '전체' },
  { value: 'draft', label: '작성중' },
  { value: 'ordered', label: '발주완료' },
  { value: 'deposit_paid', label: '선납완료' },
  { value: 'received', label: '입고완료' },
  { value: 'balance_paid', label: '잔금완료' },
];

export default function PurchasingPage() {
  const router = useRouter();
  const [statusFilter, setStatusFilter] = useState('');
  const [search, setSearch] = useState('');
  const [dateRange, setDateRange] = useState<'all' | 'today' | 'week' | 'month'>('all');
  const [page, setPage] = useState(1);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const limit = 20;

  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

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

      <div className="px-4 md:px-6 py-4 space-y-3">
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
            className="shrink-0 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-xs font-medium text-neutral-600 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
          >
            <option value="all">전체 기간</option>
            <option value="today">오늘</option>
            <option value="week">이번주</option>
            <option value="month">이번달</option>
          </select>
        </div>

        <div className="flex gap-1.5 overflow-x-auto pb-1">
          {STATUS_TABS.map((tab) => (
            <button
              key={tab.value}
              onClick={() => { setStatusFilter(tab.value); setPage(1); }}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                statusFilter === tab.value
                  ? 'bg-terracotta text-cream'
                  : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* PC: 2열 마스터-디테일 */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-240px)]">
            <div className="w-[40%] shrink-0 overflow-y-auto">
              <Card padding={false}>{listContent}</Card>
              <div className="mt-2">
                <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
              </div>
            </div>
            <div className="flex-1 min-w-0 overflow-y-auto">
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
  // 입고 전이지만 잔금까지 다 냈으면 '선납완료' 대신 '결제완료' 로 (선납완료 뱃지가 잔금 남은 것처럼 보이는 오해 방지)
  const paidEarly = !!po.balance_paid_at && (po.status === 'ordered' || po.status === 'deposit_paid');
  return (
    <div onClick={onClick} className={`flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition ${isSelected ? 'bg-terracotta/5 border-l-2 border-l-terracotta' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{po.supplier_name}</span>
          <Badge className={paidEarly ? 'bg-emerald-100 text-emerald-700' : (STATUS_COLOR[po.status] || STATUS_COLOR.draft)}>
            {paidEarly ? '결제완료' : (STATUS_LABEL[po.status] || po.status)}
          </Badge>
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
