'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useContracts } from '@/hooks/use-contracts';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { Search, Plus, ChevronLeft, ChevronRight } from 'lucide-react';
import type { Contract } from '@/lib/supabase/types';

const STATUS_TABS = [
  { value: 'all', label: '전체' },
  { value: 'draft', label: '작성중' },
  { value: 'signed', label: '서명완료' },
  { value: 'sent', label: '발송완료' },
  { value: 'completed', label: '완료' },
  { value: 'cancelled', label: '취소' },
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
  const [status, setStatus] = useState('all');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);

  const { data, isLoading } = useContracts({ status, search, page, limit: 20 });
  const contracts = data?.contracts || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  return (
    <>
      <Topbar title="전자 계약서" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="flex items-center gap-3">
          <Button size="sm" onClick={() => router.push('/contracts/new')}>
            <Plus size={14} />
            계약서 작성
          </Button>

          <div className="flex-1 relative">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => { setSearch(e.target.value); setPage(1); }}
              placeholder="계약번호, 고객명, 전화번호 검색"
              className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
            />
          </div>
        </div>

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
            </button>
          ))}
        </div>

        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : contracts.length === 0 ? (
            <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
              계약서가 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {contracts.map((c) => (
                <ContractRow key={c.id} contract={c} onClick={() => router.push(`/contracts/${c.id}`)} />
              ))}
            </div>
          )}
        </Card>

        {totalPages > 1 && (
          <div className="flex items-center justify-center gap-2">
            <Button variant="ghost" size="sm" disabled={page <= 1} onClick={() => setPage(page - 1)}>
              <ChevronLeft size={16} />
            </Button>
            <span className="text-sm text-neutral-500">{page} / {totalPages}</span>
            <Button variant="ghost" size="sm" disabled={page >= totalPages} onClick={() => setPage(page + 1)}>
              <ChevronRight size={16} />
            </Button>
          </div>
        )}
      </div>
    </>
  );
}

function ContractRow({ contract, onClick }: { contract: Contract; onClick: () => void }) {
  return (
    <div onClick={onClick} className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition">
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">
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
        <span className="text-sm font-bold">{formatKRW(contract.final_amount)}</span>
      </div>
    </div>
  );
}
