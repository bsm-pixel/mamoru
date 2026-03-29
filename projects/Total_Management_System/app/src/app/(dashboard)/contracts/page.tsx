'use client';

import { useState, useEffect, memo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useContracts, useContractTabCounts } from '@/hooks/use-contracts';
import type { ContractTab } from '@/hooks/use-contracts';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
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

export default function ContractsPage() {
  const router = useRouter();
  const [tab, setTab] = useState<ContractTab>('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [isLg, setIsLg] = useState(false);

  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1024px)');
    setIsLg(mq.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, []);

  const { data, isLoading } = useContracts({ tab, search, page, limit: 20 });
  const { data: tabCounts } = useContractTabCounts();
  const contracts = data?.contracts || [];
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

        {/* PC: 2열 마스터-디테일 */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-240px)]">
            {/* 좌측: 테이블 */}
            <div className="w-[50%] shrink-0 overflow-y-auto">
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
                      <ContractRow key={c.id} contract={c} isSelected={selectedId === c.id} onClick={() => setSelectedId(c.id)} />
                    ))}
                  </div>
                )}
              </Card>
              <div className="mt-2">
                <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
              </div>
            </div>

            {/* 우측: 상세 패널 */}
            <div className="flex-1 min-w-0 overflow-y-auto">
              {selectedId ? (
                <ContractDetailPanel contractId={selectedId} />
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
                    <ContractRow key={c.id} contract={c} isSelected={false} onClick={() => setSelectedId(c.id)} />
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
          {selectedId && <ContractDetailPanel contractId={selectedId} />}
        </SlidePanel>
      )}
    </>
  );
}

const PAYMENT_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합', cms: 'CMS',
};

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
const ContractRow = memo(function ContractRow({ contract, isSelected, onClick }: { contract: Contract; isSelected?: boolean; onClick: () => void }) {
  const isCancelled = contract.status === 'cancelled';
  return (
    <div onClick={onClick} className={`flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition ${isCancelled ? 'opacity-50' : ''} ${isSelected ? 'bg-terracotta/5 border-l-2 border-l-terracotta' : ''}`}>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className={`text-sm font-semibold text-indigo-black truncate ${isCancelled ? 'line-through' : ''}`}>
            {contract.customer_name}
          </span>
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
