'use client';

import { useState, useEffect, memo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { DeliveryDetailPanel } from '@/components/deliveries/delivery-detail-panel';
import { useDeliveries, useDeliveryStats } from '@/hooks/use-deliveries';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { Package, Plus, AlertCircle, Calendar, TrendingUp } from 'lucide-react';

/* ── 상수 ── */
const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', confirmed: '납품확정', shipped: '출고완료', settled: '정산완료',
};
const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  confirmed: 'bg-blue-100 text-blue-700',
  shipped: 'bg-green-100 text-green-700',
  settled: 'bg-emerald-100 text-emerald-700',
};
const PAYMENT_LABEL: Record<string, string> = { unpaid: '미결제', partial: '부분결제', paid: '결제완료' };
const PAYMENT_COLOR: Record<string, string> = { unpaid: 'bg-red-100 text-red-600', partial: 'bg-yellow-100 text-yellow-700', paid: 'bg-green-100 text-green-700' };
const RECEIPT_LABEL: Record<string, string> = { expense_proof: '지출증빙', tax_invoice: '세금계산서', none: '미적용' };

const STATUS_TABS = [
  { value: '', label: '전체' },
  { value: 'draft', label: '작성중' },
  { value: 'confirmed', label: '납품확정' },
  { value: 'shipped', label: '출고완료' },
];

const DATE_RANGES = [
  { value: 'all', label: '전체 기간' },
  { value: 'today', label: '오늘' },
  { value: 'week', label: '이번주' },
  { value: 'month', label: '이번달' },
] as const;

export default function DeliveriesPage() {
  const router = useRouter();
  const [statusFilter, setStatusFilter] = useState('');
  const [search, setSearch] = useState('');
  const [dateRange, setDateRange] = useState<string>('all');
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

  const { data, isLoading } = useDeliveries({
    status: statusFilter || undefined,
    search: search || undefined,
    dateRange,
    page,
    limit,
  });
  const deliveries = data?.deliveries || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);
  const { data: stats } = useDeliveryStats();

  /* ── 목록 영역 ── */
  const listContent = (
    <>
      {/* 통계 카드 */}
      {stats && (
        <div className="grid grid-cols-3 gap-2">
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <Calendar size={12} />이번주
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.weekAmount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.weekCount}건</p>
          </div>
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <TrendingUp size={12} />이번달
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.monthAmount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.monthCount}건</p>
          </div>
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-red-500 mb-1">
              <AlertCircle size={12} />미수금
            </div>
            <p className="text-base font-bold text-red-600">{formatKRW(stats.outstanding)}</p>
          </div>
        </div>
      )}

      {/* 검색 + 생성 버튼 — 2026-05-26 Phase E: 입력 진입점 /sales/new?mode=b2b 로 일원화 */}
      <div className="flex items-center gap-2">
        <Button size="sm" onClick={() => router.push('/sales/new?mode=b2b')}>
          <Plus size={14} />납품서 작성
        </Button>
        <Button size="sm" variant="secondary" onClick={() => router.push('/sales/new?mode=b2b&initial=repair')}>
          <Plus size={14} />B2B 수리
        </Button>
        <SearchInput
          value={search}
          onChange={(v) => { setSearch(v); setPage(1); }}
          placeholder="납품번호, 고객명 검색"
        />
      </div>

      {/* 탭 바 */}
      <div className="flex gap-1 border-b border-neutral-200">
        {STATUS_TABS.map((tab) => {
          const count = stats ? (tab.value === '' ? stats.all : stats[tab.value as keyof typeof stats] as number) : 0;
          const active = statusFilter === tab.value;
          return (
            <button
              key={tab.value}
              onClick={() => { setStatusFilter(tab.value); setPage(1); }}
              className={`px-3 py-2 text-sm font-medium border-b-2 transition whitespace-nowrap ${
                active
                  ? 'border-neutral-900 text-neutral-900'
                  : 'border-transparent text-neutral-400 hover:text-neutral-600'
              }`}
            >
              {tab.label}
              {typeof count === 'number' && count > 0 && (
                <span className={`ml-1.5 text-xs px-1.5 py-0.5 rounded-full ${
                  active ? 'bg-neutral-900 text-white' : 'bg-neutral-200 text-neutral-500'
                }`}>
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* 기간 필터 */}
      <div className="flex gap-1">
        {DATE_RANGES.map((d) => (
          <button
            key={d.value}
            onClick={() => { setDateRange(d.value); setPage(1); }}
            className={`px-2.5 py-1 text-xs rounded-full border transition ${
              dateRange === d.value
                ? 'bg-neutral-900 text-white border-neutral-900'
                : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-400'
            }`}
          >
            {d.label}
          </button>
        ))}
      </div>

      {/* 납품 목록 */}
      <Card padding={false}>
        {isLoading ? (
          <div className="p-4 space-y-3">
            {Array.from({ length: 5 }).map((_, i) => (
              <Skeleton key={i} className="h-14 w-full" />
            ))}
          </div>
        ) : deliveries.length === 0 ? (
          <EmptyState icon={Package} message="납품 내역이 없습니다" />
        ) : (
          <div className="divide-y divide-neutral-100">
            {deliveries.map((dl) => (
              <DeliveryRow
                key={dl.id as string}
                dl={dl}
                isSelected={selectedId === dl.id}
                onClick={() => setSelectedId(dl.id as string)}
              />
            ))}
          </div>
        )}
      </Card>

      <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
    </>
  );

  return (
    <>
      <Topbar title="B2B거래" />

      {isLg ? (
        /* PC: 마스터-디테일 2컬럼 */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-full min-h-0">
          <div className="w-2/5 shrink-0 overflow-y-auto space-y-3 pr-1">
            {listContent}
          </div>
          <div className="flex-1 min-w-0 overflow-y-auto bg-white rounded-xl border border-neutral-200">
            {selectedId ? (
              <DeliveryDetailPanel deliveryId={selectedId} />
            ) : (
              <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                <Package size={28} className="mb-2 opacity-40" />
                <p className="text-xs">목록에서 납품 건을 선택하세요</p>
              </div>
            )}
          </div>
        </div>
      ) : (
        /* 모바일: 목록 */
        <div className="px-4 md:px-6 py-4 space-y-4">
          {listContent}
        </div>
      )}

      {/* 모바일 상세 패널 */}
      {!isLg && (
        <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="납품 상세" className="sm:w-[480px]">
          {selectedId && <DeliveryDetailPanel deliveryId={selectedId} />}
        </SlidePanel>
      )}

      {/* 2026-05-26 Phase E: 납품서 작성 모달은 /sales/new?mode=b2b 페이지에서 처리 (IA 통합) */}
    </>
  );
}

/* ── 목록 행 ── */
const DeliveryRow = memo(function DeliveryRow({ dl, isSelected, onClick }: {
  dl: Record<string, unknown>; isSelected: boolean; onClick: () => void;
}) {
  const status = (dl.status as string) || 'draft';
  const paymentStatus = (dl.payment_status as string) || 'unpaid';

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer transition ${
        isSelected ? 'bg-neutral-50 border-l-2 border-l-neutral-900' : 'hover:bg-warm-ivory/60'
      }`}
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-sm font-semibold text-indigo-black truncate">
            {dl.customer_name as string}
          </span>
          <Badge className={STATUS_COLOR[status === 'settled' ? 'shipped' : status] || STATUS_COLOR.draft}>
            {status === 'settled' ? '출고완료' : (STATUS_LABEL[status] || status)}
          </Badge>
          <Badge className={PAYMENT_COLOR[paymentStatus] || PAYMENT_COLOR.unpaid}>
            {PAYMENT_LABEL[paymentStatus] || paymentStatus}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{formatDate(dl.delivery_date as string)}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        {dl.payment_status === 'partial' && (dl.paid_amount as number) > 0 && (
          <p className="text-[10px] text-yellow-600 font-medium">{formatKRW((dl.paid_amount as number))} 선납</p>
        )}
        <p className="text-sm font-bold">{formatKRW((dl.total_amount as number) || 0)}</p>
      </div>
    </div>
  );
});

