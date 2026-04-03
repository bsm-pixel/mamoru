'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomers } from '@/hooks/use-customers';
import { CustomerCreateModal } from '@/components/customers/customer-create-modal';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { CustomerDetailPanel } from '@/components/customers/customer-detail-panel';
import { Users, Plus, X } from 'lucide-react';
import type { Customer } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

const TYPE_LABEL: Record<string, string> = {
  retail: '일반',
  online: '온라인',
  dealer: '딜러',
  academy: '아카데미',
};

const TYPE_COLOR: Record<string, string> = {
  retail: 'bg-neutral-100 text-neutral-600',
  online: 'bg-blue-100 text-blue-700',
  dealer: 'bg-purple-100 text-purple-700',
  academy: 'bg-emerald-100 text-emerald-700',
};

const SOURCE_LABEL: Record<string, string> = {
  imweb: '아임웹',
  consultation: '상담',
  as: '복원수리',
  manual: '수동',
};

const FILTER_TYPES = [
  { value: '', label: '전체' },
  { value: 'retail', label: '일반' },
  { value: 'online', label: '온라인' },
  { value: 'dealer', label: '딜러' },
  { value: 'academy', label: '아카데미' },
];

export default function CustomersPage() {
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('');
  const [page, setPage] = useState(1);
  const [showAdd, setShowAdd] = useState(false);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const limit = 20;

  // PC 여부 감지
  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  const { data, isLoading } = useCustomers({
    search,
    type: typeFilter || undefined,
    page,
    limit,
  });
  const customers = data?.customers || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);

  const listContent = isLoading ? (
    <div className="p-4 space-y-3">
      {Array.from({ length: 5 }).map((_, i) => (
        <Skeleton key={i} className="h-16 w-full" />
      ))}
    </div>
  ) : customers.length === 0 ? (
    <EmptyState icon={Users} message="고객이 없습니다" />
  ) : (
    <div className="divide-y divide-neutral-100">
      {customers.map((c) => (
        <CustomerRow
          key={c.id}
          customer={c}
          isSelected={selectedId === c.id}
          onClick={() => setSelectedId(c.id)}
        />
      ))}
    </div>
  );

  return (
    <>
      <Topbar title="고객 관리" action={
        <Button size="sm" onClick={() => setShowAdd(true)}>
          <Plus size={14} />
          고객 추가
        </Button>
      } />

      <div className="px-4 md:px-6 py-4 space-y-3">
        {/* 검색 */}
        <SearchInput
          value={search}
          onChange={(v) => { setSearch(v); setPage(1); }}
          placeholder="고객명, 전화번호, 업체명 검색"
        />

        {/* 유형 필터 칩 */}
        <div className="flex gap-1.5 overflow-x-auto pb-1">
          {FILTER_TYPES.map((f) => (
            <button
              key={f.value}
              onClick={() => { setTypeFilter(f.value); setPage(1); }}
              className={`shrink-0 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                typeFilter === f.value
                  ? 'bg-terracotta text-cream'
                  : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}
            >
              {f.label}
            </button>
          ))}
        </div>

        {/* PC: 2열 레이아웃 */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-240px)]">
            {/* 좌측: 고객 목록 */}
            <div className="w-[40%] shrink-0 overflow-y-auto">
              <Card padding={false}>{listContent}</Card>
              <div className="mt-2">
                <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} unit="명" />
              </div>
            </div>

            {/* 우측: 고객 상세 모니터 */}
            <div className="flex-1 min-w-0 overflow-y-auto">
              {selectedId ? (
                <CustomerDetailPanel customerId={selectedId} />
              ) : (
                <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                  <Users size={28} className="mb-2 opacity-40" />
                  <p className="text-xs">목록에서 고객을 선택하세요</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* 모바일: 목록만 */}
        {!isLg && (
          <>
            <Card padding={false}>{listContent}</Card>
            <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} unit="명" />
          </>
        )}
      </div>

      <CustomerCreateModal
        open={showAdd}
        onClose={() => setShowAdd(false)}
        onCreated={() => { /* 목록 자동 갱신 (queryClient invalidation은 모달 내부에서 처리) */ }}
      />

      {/* 모바일 전용 슬라이드 패널 */}
      {!isLg && (
        <SlidePanel
          open={!!selectedId}
          onClose={() => setSelectedId(null)}
          title="고객 상세"
          className="sm:w-[640px]"
        >
          {selectedId && <CustomerDetailPanel customerId={selectedId} />}
        </SlidePanel>
      )}
    </>
  );
}

/* 기존 AddCustomerModal → CustomerCreateModal 공통 컴포넌트로 교체 (customer-create-modal.tsx) */

function CustomerRow({ customer, isSelected, onClick }: { customer: Customer; isSelected: boolean; onClick: () => void }) {
  const c = customer;
  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition ${isSelected ? 'bg-terracotta/5 border-l-2 border-l-terracotta' : ''}`}
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
          <Badge className={TYPE_COLOR[c.customer_type] || TYPE_COLOR.retail}>
            {TYPE_LABEL[c.customer_type] || '일반'}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          {c.phone && <span>{formatPhone(c.phone)}</span>}
          {c.company_name && <span>{c.company_name}</span>}
          <span>{SOURCE_LABEL[c.source] || c.source}</span>
          <span>{formatDate(c.created_at)}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        {c.total_spent > 0 && (
          <p className="text-sm font-bold">{formatKRW(c.total_spent)}</p>
        )}
        {c.outstanding_balance > 0 && (
          <p className="text-xs text-red-500 font-semibold">미수 {formatKRW(c.outstanding_balance)}</p>
        )}
      </div>
    </div>
  );
}
