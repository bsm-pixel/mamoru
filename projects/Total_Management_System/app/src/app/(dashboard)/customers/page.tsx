'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useCustomers } from '@/hooks/use-customers';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { Search, ChevronLeft, ChevronRight } from 'lucide-react';
import type { Customer } from '@/lib/supabase/types';

const TYPE_LABEL: Record<string, string> = {
  retail: '일반',
  online: '온라인',
  dealer: '딜러',
  supplier: '매입처',
};

const TYPE_COLOR: Record<string, string> = {
  retail: 'bg-neutral-100 text-neutral-600',
  online: 'bg-blue-100 text-blue-700',
  dealer: 'bg-purple-100 text-purple-700',
  supplier: 'bg-amber-100 text-amber-700',
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
  { value: 'supplier', label: '매입처' },
];

export default function CustomersPage() {
  const router = useRouter();
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('');
  const [page, setPage] = useState(1);
  const limit = 20;

  const { data, isLoading } = useCustomers({
    search,
    type: typeFilter || undefined,
    page,
    limit,
  });
  const customers = data?.customers || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);

  return (
    <>
      <Topbar title="고객 관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 검색 */}
        <div className="relative">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => { setSearch(e.target.value); setPage(1); }}
            placeholder="고객명, 전화번호, 업체명 검색"
            className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
          />
        </div>

        {/* 유형 필터 */}
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

        {/* 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : customers.length === 0 ? (
            <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
              고객이 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {customers.map((c) => (
                <CustomerRow
                  key={c.id}
                  customer={c}
                  onClick={() => router.push(`/customers/${c.id}`)}
                />
              ))}
            </div>
          )}
        </Card>

        {/* 건수 + 페이지네이션 */}
        <div className="flex items-center justify-between">
          <span className="text-xs text-neutral-400">총 {total}명</span>
          {totalPages > 1 && (
            <div className="flex items-center gap-2">
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
      </div>
    </>
  );
}

function CustomerRow({ customer, onClick }: { customer: Customer; onClick: () => void }) {
  const c = customer;
  return (
    <div
      onClick={onClick}
      className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{c.name}</span>
          <Badge className={TYPE_COLOR[c.customer_type] || TYPE_COLOR.retail}>
            {TYPE_LABEL[c.customer_type] || '일반'}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          {c.phone && <span>{c.phone}</span>}
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
