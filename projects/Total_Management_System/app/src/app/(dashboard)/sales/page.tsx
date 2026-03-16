'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSales } from '@/hooks/use-sales';
import { useContracts } from '@/hooks/use-contracts';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { Plus, FileSignature, Receipt } from 'lucide-react';
import type { OfflineSale } from '@/lib/supabase/types';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드',
  cash: '현금',
  transfer: '계좌이체',
  mixed: '복합',
};

const PAYMENT_STATUS_COLOR: Record<string, string> = {
  paid: 'bg-green-100 text-green-700',
  unpaid: 'bg-red-100 text-red-700',
  partial: 'bg-yellow-100 text-yellow-700',
};

export default function SalesPage() {
  const router = useRouter();
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);

  const { data, isLoading } = useSales({ search, page, limit: 20 });
  const { data: contractData } = useContracts({ status: 'signed', limit: 100 });
  const newContractCount = contractData?.contracts?.filter((c) => !c.offline_sale_id).length || 0;
  const sales = data?.sales || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / 20);

  return (
    <>
      <Topbar title="판매관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 신규 계약서 알림 */}
        {newContractCount > 0 && (
          <button
            onClick={() => router.push('/contracts')}
            className="w-full flex items-center gap-3 px-4 py-3 rounded-lg bg-terracotta/10 border border-terracotta/20 hover:bg-terracotta/15 transition text-left"
          >
            <FileSignature size={18} className="text-terracotta shrink-0" />
            <div className="flex-1">
              <p className="text-sm font-semibold text-terracotta">신규 계약서 {newContractCount}건</p>
              <p className="text-xs text-neutral-500">매칭 대기 중인 계약서가 있습니다</p>
            </div>
          </button>
        )}

        {/* 상단: 신규 판매 + 검색 */}
        <div className="flex items-center gap-3">
          <Button
            size="sm"
            onClick={() => router.push('/sales/new')}
          >
            <Plus size={14} />
            판매 입력
          </Button>

          <SearchInput
            value={search}
            onChange={(v) => { setSearch(v); setPage(1); }}
            placeholder="판매번호, 고객명, 전화번호 검색"
          />
        </div>

        {/* 판매 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-3">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-16 w-full" />
              ))}
            </div>
          ) : sales.length === 0 ? (
            <EmptyState icon={Receipt} message="판매 기록이 없습니다" />
          ) : (
            <div className="divide-y divide-neutral-100">
              {sales.map((sale) => (
                <SaleRow
                  key={sale.id}
                  sale={sale}
                  onClick={() => router.push(`/sales/${sale.id}`)}
                />
              ))}
            </div>
          )}
        </Card>

        <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
      </div>
    </>
  );
}

function SaleRow({ sale, onClick }: { sale: OfflineSale; onClick: () => void }) {
  return (
    <div
      onClick={onClick}
      className="flex items-center gap-4 px-4 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black truncate">
            {sale.customer_name}
          </span>
          <Badge className={PAYMENT_STATUS_COLOR[sale.payment_status] || ''}>
            {sale.payment_status === 'paid' ? '결제완료' : sale.payment_status === 'unpaid' ? '미결제' : '부분결제'}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{sale.sale_number}</span>
          <span>{formatDate(sale.sale_date)}</span>
          <span>{PAYMENT_METHOD_LABEL[sale.payment_method] || sale.payment_method}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        <span className="text-sm font-bold">{formatKRW(sale.paid_amount)}</span>
      </div>
    </div>
  );
}
