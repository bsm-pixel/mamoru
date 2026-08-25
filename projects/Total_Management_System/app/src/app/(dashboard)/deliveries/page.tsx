'use client';

import { useState, memo, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { getDeliveryStatusChip, type DeliveryStatusInput } from '@/lib/deliveries/status';
import { getDeliveryShipStatus } from '@/lib/sales/ship-status';
import { deliveryNet } from '@/lib/sales/amounts';
import { SlidePanel } from '@/components/ui/slide-panel';
import { DataGrid, type GridColumn } from '@/components/ui/data-grid';
import { DeliveryDetailPanel } from '@/components/deliveries/delivery-detail-panel';
import { useActivityTypes, type ActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';
import { useDeliveries, useDeliveryStats } from '@/hooks/use-deliveries';
import { useIsLg } from '@/hooks/use-grid-mode';
import { formatKRW, formatDate, CUSTOMER_TYPE_LABEL, CUSTOMER_TYPE_COLOR } from '@/lib/utils/format';
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
// 고객유형 라벨/색은 format.ts SSOT(CUSTOMER_TYPE_LABEL/COLOR) 사용
/** 배송상태 tone → 텍스트 색 (sales SHIP_TONE 과 동일) */
const SHIP_TONE = { amber: 'text-amber-600', green: 'text-emerald-600', mute: 'text-neutral-300' } as const;

/** 그리드 행 타입 — deliveries 는 훅에서 Record<string,unknown> 로 오므로 표시용 필드만 명시 */
interface DeliveryLike extends DeliveryStatusInput {
  id: string;
  customer_name?: string | null;
  company_name?: string | null;
  customer_type?: string | null;
  delivery_date?: string | null;
  payment_status?: string | null;
  total_amount?: number | null;
  discount_amount?: number | null;
}

/** PC 그리드 컬럼: 납품일·거래처·유형·상태·배송·결제·금액 */
const DELIVERY_COLUMNS: GridColumn<DeliveryLike>[] = [
  { key: 'date', label: '납품일', render: (d) => <span className="text-neutral-600 whitespace-nowrap tabular-nums">{d.delivery_date ? formatDate(d.delivery_date) : '—'}</span> },
  { key: 'customer', label: '거래처', render: (d) => <span className="font-semibold text-indigo-black truncate">{d.company_name || d.customer_name || '미지정'}</span> },
  { key: 'type', label: '유형', render: (d) => <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${CUSTOMER_TYPE_COLOR[d.customer_type || ''] || 'bg-neutral-100 text-neutral-500'}`}>{CUSTOMER_TYPE_LABEL[d.customer_type || ''] || '거래처'}</span> },
  { key: 'status', label: '상태', render: (d) => { const c = getDeliveryStatusChip(d); return <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${c.className}`}>{c.label}</span>; } },
  { key: 'ship', label: '배송', render: (d) => { const s = getDeliveryShipStatus(d); return <span className={`text-xs font-medium ${SHIP_TONE[s.tone]}`}>{s.label}</span>; } },
  { key: 'pay', label: '결제', render: (d) => <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${PAYMENT_COLOR[d.payment_status || ''] || PAYMENT_COLOR.unpaid}`}>{PAYMENT_LABEL[d.payment_status || ''] || '미결제'}</span> },
  { key: 'amount', label: '금액', align: 'right', render: (d) => <span className="font-bold tabular-nums text-indigo-black">{formatKRW(deliveryNet(d))}</span> },
];

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

  const isLg = useIsLg();

  const { data, isLoading } = useDeliveries({
    status: statusFilter || undefined,
    search: search || undefined,
    dateRange,
    page,
    limit,
  });
  const deliveries = data?.deliveries || [];
  const deliveryActTypes = useActivityTypes(deliveries.map((d) => (d as { customer_phone?: string | null }).customer_phone));
  const columns = useMemo<GridColumn<DeliveryLike>[]>(() =>
    DELIVERY_COLUMNS.map((col) => col.key === 'customer'
      ? { ...col, render: (d: DeliveryLike) => (
          <span className="inline-flex items-center gap-1 min-w-0 max-w-full">
            <span className="font-semibold text-indigo-black truncate">{d.company_name || d.customer_name || '미지정'}</span>
            <ActivityChips types={deliveryActTypes((d as { customer_phone?: string | null }).customer_phone)} className="shrink-0" />
          </span>
        ) }
      : col
    ), [deliveryActTypes]);
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);
  const { data: stats } = useDeliveryStats();

  /* ── 목록 영역 ── */
  const listContent = (
    <>
      {/* 통계 카드 */}
      {stats && (
        <div className="grid grid-cols-3 gap-2">
          <div className="bg-white rounded-2xl border border-stone-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <Calendar size={12} />이번주
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.weekAmount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.weekCount}건</p>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <TrendingUp size={12} />이번달
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.monthAmount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.monthCount}건</p>
          </div>
          <div className="bg-white rounded-2xl border border-stone-200 p-3">
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
        ) : isLg ? (
          <DataGrid
            columns={columns}
            rows={deliveries as unknown as DeliveryLike[]}
            getRowKey={(d) => d.id}
            selectedKey={selectedId ?? undefined}
            onSelect={(d) => setSelectedId(d.id)}
            rowClassName={(d) => (d.cancelled_at ? 'opacity-50' : '')}
            emptyMessage="납품 내역이 없습니다"
          />
        ) : (
          <div className="divide-y divide-neutral-100">
            {deliveries.map((dl) => (
              <DeliveryRow
                key={dl.id as string}
                dl={dl}
                isSelected={selectedId === dl.id}
                onClick={() => setSelectedId(dl.id as string)}
                actTypes={deliveryActTypes((dl as { customer_phone?: string | null }).customer_phone)}
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
      <Topbar title="B2B 납품" />

      {isLg ? (
        /* PC: 밀집 그리드 + 우측 상세 (카드보기·토글 폐지 — 항상 그리드) */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-full min-h-0 bg-stone-50">
          <div className="flex-1 min-w-0 overflow-auto space-y-3 pr-1">
            {listContent}
          </div>
          <div className="w-[440px] shrink-0 overflow-y-auto bg-white rounded-2xl border border-stone-200">
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
        <div className="px-4 md:px-6 py-4 space-y-4 bg-stone-50 min-h-screen">
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
const DeliveryRow = memo(function DeliveryRow({ dl, isSelected, onClick, actTypes }: {
  dl: Record<string, unknown>; isSelected: boolean; onClick: () => void; actTypes?: ActivityTypes;
}) {
  const status = (dl.status as string) || 'draft';
  const paymentStatus = (dl.payment_status as string) || 'unpaid';

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer transition ${
        isSelected ? 'bg-stone-50 border-l-2 border-l-stone-900' : 'hover:bg-stone-50'
      }`}
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-sm font-semibold text-stone-900 truncate">
            {dl.customer_name as string}
          </span>
          <ActivityChips types={actTypes} className="shrink-0" />
          {/* 110: 상세와 같은 규칙 (납품확정 → 출고대기 → 출고완료 → 배송완료) */}
          {(() => {
            const chip = getDeliveryStatusChip(dl);
            return <Badge className={chip.className}>{chip.label}</Badge>;
          })()}
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

