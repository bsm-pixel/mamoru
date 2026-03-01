'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { usePurchaseOrders } from '@/hooks/use-purchasing';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { Search, Plus, ChevronLeft, ChevronRight } from 'lucide-react';
import type { PurchaseOrder } from '@/lib/supabase/types';

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중',
  ordered: '발주완료',
  deposit_paid: '선납완료',
  received: '입고완료',
  balance_paid: '잔금완료',
  cancelled: '취소',
};

const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  ordered: 'bg-blue-100 text-blue-700',
  deposit_paid: 'bg-yellow-100 text-yellow-700',
  received: 'bg-green-100 text-green-700',
  balance_paid: 'bg-emerald-100 text-emerald-700',
  cancelled: 'bg-red-100 text-red-700',
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
  const [page, setPage] = useState(1);
  const limit = 20;

  const { data, isLoading } = usePurchaseOrders({
    status: statusFilter || undefined,
    search: search || undefined,
    page,
    limit,
  });
  const orders = data?.orders || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);

  return (
    <>
      <Topbar title="매입관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="flex items-center gap-3">
          <Button size="sm" onClick={() => router.push('/purchasing/new')}>
            <Plus size={14} />
            발주 작성
          </Button>
          <div className="flex-1 relative">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => { setSearch(e.target.value); setPage(1); }}
              placeholder="발주번호, 매입처 검색"
              className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
            />
          </div>
        </div>

        {/* 상태 탭 */}
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

        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : orders.length === 0 ? (
            <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
              발주 내역이 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {orders.map((po) => (
                <PORow key={po.id} po={po} onClick={() => router.push(`/purchasing/${po.id}`)} />
              ))}
            </div>
          )}
        </Card>

        <div className="flex items-center justify-between">
          <span className="text-xs text-neutral-400">총 {total}건</span>
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

function PORow({ po, onClick }: { po: PurchaseOrder; onClick: () => void }) {
  return (
    <div onClick={onClick} className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition">
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">{po.supplier_name}</span>
          <Badge className={STATUS_COLOR[po.status] || STATUS_COLOR.draft}>
            {STATUS_LABEL[po.status] || po.status}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{po.po_number}</span>
          <span>{formatDate(po.order_date)}</span>
          {po.expected_date && <span>예정: {formatDate(po.expected_date)}</span>}
        </div>
      </div>
      <div className="text-right shrink-0">
        <p className="text-sm font-bold">{formatKRW(po.total_amount)}</p>
        {po.deposit_amount > 0 && po.status !== 'balance_paid' && (
          <p className="text-xs text-neutral-500">선납 {formatKRW(po.deposit_amount)}</p>
        )}
      </div>
    </div>
  );
}
