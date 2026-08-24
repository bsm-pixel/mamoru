'use client';

import { useState, memo, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useContracts, useContractTabCounts } from '@/hooks/use-contracts';
import type { ContractTab } from '@/hooks/use-contracts';
import { useIsLg } from '@/hooks/use-grid-mode';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { DataGrid, type GridColumn } from '@/components/ui/data-grid';
import { useActivityTypes, type ActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';
import { ContractDetailPanel } from '@/components/contracts/contract-detail-panel';
import { Plus, FileText } from 'lucide-react';
import type { Contract } from '@/lib/supabase/types';

const TABS: { key: ContractTab; label: string }[] = [
  { key: 'all', label: '전체' },
  { key: 'new', label: '신규계약' },
  { key: 'converted', label: '전환완료' },
  { key: 'cancelled', label: '취소' },
];

const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  signed: 'bg-blue-100 text-blue-700',
  sent: 'bg-purple-100 text-purple-700',
  completed: 'bg-green-100 text-green-700',
  cancelled: 'bg-red-100 text-red-700',
};

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중',
  signed: '서명완료',
  sent: '발송완료',
  completed: '완료',
  cancelled: '취소',
};

const PAYMENT_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합', cms: 'CMS',
};

/** PC 그리드 컬럼: 계약번호·작성일·고객·상태·결제·금액 */
const CONTRACT_COLUMNS: GridColumn<Contract>[] = [
  { key: 'no', label: '계약번호', render: (c) => <span className={`font-mono text-[11px] text-neutral-500 ${c.status === 'cancelled' ? 'line-through' : ''}`}>{c.contract_number}</span> },
  { key: 'created', label: '작성일', render: (c) => <span className="text-neutral-600 whitespace-nowrap tabular-nums">{formatDate(c.created_at)}</span> },
  { key: 'customer', label: '고객', render: (c) => <span className={`font-semibold text-indigo-black truncate ${c.status === 'cancelled' ? 'line-through' : ''}`}>{c.customer_name}</span> },
  { key: 'status', label: '상태', render: (c) => <span className={`px-2 py-0.5 rounded text-[10.5px] font-bold ${STATUS_COLOR[c.status] || 'bg-neutral-100 text-neutral-500'}`}>{STATUS_LABEL[c.status] || c.status}</span> },
  { key: 'pay', label: '결제', render: (c) => <span className="text-neutral-500 text-xs">{PAYMENT_LABEL[c.payment_method] || c.payment_method || '—'}</span> },
  { key: 'amount', label: '금액', align: 'right', render: (c) => <span className={`font-bold tabular-nums ${c.status === 'cancelled' ? 'line-through text-neutral-400' : 'text-indigo-black'}`}>{formatKRW(c.final_amount)}</span> },
];

export default function ContractsPage() {
  const router = useRouter();
  const [tab, setTab] = useState<ContractTab>('all');
  const [search, setSearch] = useState('');
  const [dateRange, setDateRange] = useState<'all' | 'today' | 'week' | 'month'>('all');
  const [page, setPage] = useState(1);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const isLg = useIsLg();

  const { data, isLoading } = useContracts({ tab, search, dateRange, page, limit: 20 });
  const { data: tabCounts } = useContractTabCounts();
  const contracts = data?.contracts || [];
  const contractActTypes = useActivityTypes(contracts.map((c) => c.customer_phone));
  const columns = useMemo<GridColumn<Contract>[]>(() =>
    CONTRACT_COLUMNS.map((col) => col.key === 'customer'
      ? { ...col, render: (c: Contract) => (
          <span className="inline-flex items-center gap-1 min-w-0 max-w-full">
            <span className={`font-semibold text-indigo-black truncate ${c.status === 'cancelled' ? 'line-through' : ''}`}>{c.customer_name}</span>
            <ActivityChips types={contractActTypes(c.customer_phone)} className="shrink-0" />
          </span>
        ) }
      : col
    ), [contractActTypes]);
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  const handleTabChange = (t: ContractTab) => { setTab(t); setPage(1); };

  return (
    <>
      <Topbar title="전자 계약서" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="flex items-center gap-3">
          <Button size="sm" onClick={() => router.push('/contracts/new')}>
            <Plus size={14} />
            계약서 작성
          </Button>

          <SearchInput
            value={search}
            onChange={(v) => { setSearch(v); setPage(1); }}
            placeholder="계약번호, 고객명, 전화번호 검색"
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
        </div>

        {/* 탭 바 — 판매관리와 동일 패턴 */}
        <div className="flex gap-1 border-b border-neutral-200">
          {TABS.map((t) => {
            const count = tabCounts?.[t.key] ?? 0;
            const active = tab === t.key;
            return (
              <button
                key={t.key}
                onClick={() => handleTabChange(t.key)}
                className={`px-3 py-2 text-sm font-medium border-b-2 transition whitespace-nowrap ${
                  active
                    ? 'border-neutral-900 text-neutral-900'
                    : 'border-transparent text-neutral-400 hover:text-neutral-600'
                }`}
              >
                {t.label}
                {count > 0 && (
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

        {/* PC: 밀집 그리드 + 우측 상세 (카드보기·토글 폐지 — 항상 그리드) */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-240px)]">
            {/* 좌측: 밀집 그리드 목록 */}
            <div className="flex-1 min-w-0 overflow-auto">
              <Card padding={false}>
                {isLoading ? (
                  <div className="p-4 space-y-3">
                    {Array.from({ length: 5 }).map((_, i) => (
                      <Skeleton key={i} className="h-16 w-full" />
                    ))}
                  </div>
                ) : (
                  <DataGrid
                    columns={columns}
                    rows={contracts}
                    getRowKey={(c) => c.id}
                    selectedKey={selectedId ?? undefined}
                    onSelect={(c) => setSelectedId(c.id)}
                    rowClassName={(c) => (c.status === 'cancelled' ? 'opacity-50' : '')}
                    emptyMessage="계약서가 없습니다"
                  />
                )}
              </Card>
              <div className="mt-2">
                <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
              </div>
            </div>

            {/* 우측: 상세 패널 (고정폭) */}
            <div className="w-[440px] shrink-0 overflow-y-auto">
              {selectedId ? (
                <ContractDetailPanel contractId={selectedId} onDeleted={() => setSelectedId(null)} />
              ) : (
                <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                  <FileText size={28} className="mb-2 opacity-40" />
                  <p className="text-xs">목록에서 계약서를 선택하세요</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* 모바일: 카드 뷰 */}
        {!isLg && (
          <>
            <Card padding={false}>
              {isLoading ? (
                <div className="p-4 space-y-3">
                  {Array.from({ length: 5 }).map((_, i) => (
                    <Skeleton key={i} className="h-16 w-full" />
                  ))}
                </div>
              ) : contracts.length === 0 ? (
                <EmptyState icon={FileText} message="계약서가 없습니다" />
              ) : (
                <div className="divide-y divide-neutral-100">
                  {contracts.map((c) => (
                    <ContractRow key={c.id} contract={c} isSelected={false} onClick={() => setSelectedId(c.id)} actTypes={contractActTypes(c.customer_phone)} />
                  ))}
                </div>
              )}
            </Card>
            <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
          </>
        )}
      </div>

      {/* 모바일 전용 슬라이드 패널 */}
      {!isLg && (
        <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="계약서 상세" className="sm:w-[640px]">
          {selectedId && <ContractDetailPanel contractId={selectedId} onDeleted={() => setSelectedId(null)} />}
        </SlidePanel>
      )}
    </>
  );
}

/* PC 테이블 행 */
const ContractTableRow = memo(function ContractTableRow({ contract, onClick }: { contract: Contract; onClick: () => void }) {
  const isCancelled = contract.status === 'cancelled';
  return (
    <tr onClick={onClick} className={`border-b border-neutral-100 cursor-pointer hover:bg-neutral-50 transition ${isCancelled ? 'opacity-50' : ''}`}>
      <td className={`px-4 py-3 font-mono text-xs ${isCancelled ? 'line-through' : ''}`}>{contract.contract_number}</td>
      <td className="px-4 py-3 text-neutral-600">{formatDate(contract.created_at)}</td>
      <td className={`px-4 py-3 font-semibold ${isCancelled ? 'line-through' : ''}`}>{contract.customer_name}</td>
      <td className="px-4 py-3">
        <Badge className={STATUS_COLOR[contract.status] || ''}>{STATUS_LABEL[contract.status] || contract.status}</Badge>
      </td>
      <td className="px-4 py-3 text-neutral-600">{PAYMENT_LABEL[contract.payment_method] || contract.payment_method}</td>
      <td className={`px-4 py-3 text-right font-bold ${isCancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(contract.final_amount)}</td>
    </tr>
  );
});

/* 모바일 카드 행 */
const ContractRow = memo(function ContractRow({ contract, isSelected, onClick, actTypes }: { contract: Contract; isSelected?: boolean; onClick: () => void; actTypes?: ActivityTypes }) {
  const isCancelled = contract.status === 'cancelled';
  return (
    <div onClick={onClick} className={`flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-stone-50/60 transition ${isCancelled ? 'opacity-50' : ''} ${isSelected ? 'bg-stone-100 border-l-2 border-l-stone-900' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className={`text-sm font-semibold text-stone-900 truncate ${isCancelled ? 'line-through' : ''}`}>
            {contract.customer_name}
          </span>
          <ActivityChips types={actTypes} className="shrink-0" />
          <Badge className={STATUS_COLOR[contract.status] || ''}>
            {STATUS_LABEL[contract.status] || contract.status}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{contract.contract_number}</span>
          <span>{formatDate(contract.created_at)}</span>
          {contract.signed_at && <span className="text-blue-600">서명완료</span>}
        </div>
      </div>
      <div className="text-right shrink-0">
        <span className={`text-sm font-bold ${isCancelled ? 'line-through text-neutral-400' : ''}`}>{formatKRW(contract.final_amount)}</span>
      </div>
    </div>
  );
});
