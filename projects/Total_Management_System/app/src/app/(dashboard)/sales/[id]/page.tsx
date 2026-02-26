'use client';

import { use } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useEcountSync } from '@/hooks/use-sales';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { ArrowLeft, RefreshCw } from 'lucide-react';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드',
  cash: '현금',
  transfer: '계좌이체',
  mixed: '복합',
};

export default function SaleDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();
  const { data, isLoading } = useSale(id);
  const ecountSync = useEcountSync();

  if (isLoading) {
    return (
      <>
        <Topbar title="판매 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-48" />
          <Skeleton className="h-32" />
        </div>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <Topbar title="판매 상세" />
        <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
          판매 정보를 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { sale, items } = data;

  return (
    <>
      <Topbar title="판매 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <Button variant="ghost" size="sm" onClick={() => router.push('/sales')}>
          <ArrowLeft size={14} />
          목록으로
        </Button>

        {/* 판매 정보 */}
        <Card>
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-sm font-bold text-indigo-black">{sale.sale_number}</h3>
            <div className="flex items-center gap-2">
              <Badge className={
                sale.payment_status === 'paid' ? 'bg-green-100 text-green-700'
                  : sale.payment_status === 'unpaid' ? 'bg-red-100 text-red-700'
                  : 'bg-yellow-100 text-yellow-700'
              }>
                {sale.payment_status === 'paid' ? '결제완료' : sale.payment_status === 'unpaid' ? '미결제' : '부분결제'}
              </Badge>
              <Badge className={
                sale.ecount_sync_status === 'synced' ? 'bg-blue-100 text-blue-700'
                  : sale.ecount_sync_status === 'failed' ? 'bg-red-100 text-red-700'
                  : 'bg-neutral-100 text-neutral-500'
              }>
                {sale.ecount_sync_status === 'synced' ? 'ERP 연동' : sale.ecount_sync_status === 'failed' ? 'ERP 실패' : 'ERP 대기'}
              </Badge>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-3 text-sm">
            <div>
              <span className="text-xs text-neutral-500">고객명</span>
              <p className="font-semibold">{sale.customer_name}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">연락처</span>
              <p>{sale.customer_phone || '-'}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">판매일</span>
              <p>{formatDate(sale.sale_date, 'yyyy.MM.dd')}</p>
            </div>
            <div>
              <span className="text-xs text-neutral-500">결제방법</span>
              <p>{PAYMENT_METHOD_LABEL[sale.payment_method] || sale.payment_method}</p>
            </div>
          </div>

          {sale.memo && (
            <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">
              {sale.memo}
            </p>
          )}
        </Card>

        {/* 판매 항목 */}
        <Card>
          <h3 className="text-sm font-bold text-indigo-black mb-3">판매 항목</h3>
          <div className="space-y-2">
            {items.map((item) => (
              <div key={item.id} className="flex items-center justify-between py-2 border-b border-neutral-50 last:border-0">
                <div>
                  <p className="text-sm font-medium">{item.product_name}</p>
                  <p className="text-xs text-neutral-500">
                    {item.sku && `${item.sku} · `}{formatKRW(item.unit_price)} x {item.quantity}
                  </p>
                </div>
                <span className="text-sm font-bold">{formatKRW(item.total_price)}</span>
              </div>
            ))}
          </div>

          <div className="mt-3 pt-3 border-t border-neutral-200 space-y-1">
            <div className="flex justify-between text-sm">
              <span className="text-neutral-500">소계</span>
              <span>{formatKRW(sale.total_amount)}</span>
            </div>
            {sale.discount_amount > 0 && (
              <div className="flex justify-between text-sm">
                <span className="text-neutral-500">할인</span>
                <span className="text-red-600">-{formatKRW(sale.discount_amount)}</span>
              </div>
            )}
            <div className="flex justify-between text-sm font-bold">
              <span>결제 금액</span>
              <span className="text-terracotta">{formatKRW(sale.paid_amount)}</span>
            </div>
          </div>
        </Card>

        {/* 이카운트 연동 */}
        {sale.ecount_sync_status !== 'synced' && (
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3">이카운트 ERP 연동</h3>
            <p className="text-xs text-neutral-500 mb-3">
              판매 전표를 이카운트 ERP에 동기화합니다.
              {sale.ecount_sync_status === 'failed' && ' (이전 동기화 실패 — 재시도 가능)'}
            </p>
            <Button
              variant="secondary"
              size="sm"
              onClick={() => ecountSync.mutate(sale.id)}
              disabled={ecountSync.isPending}
            >
              <RefreshCw size={14} className={ecountSync.isPending ? 'animate-spin' : ''} />
              {ecountSync.isPending ? '동기화 중...' : '이카운트 동기화'}
            </Button>
          </Card>
        )}

        {sale.ecount_sync_status === 'synced' && sale.ecount_slip_no && (
          <Card>
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-indigo-black">이카운트 ERP 연동</h3>
                <p className="text-xs text-neutral-500 mt-1">
                  전표번호: {sale.ecount_slip_no}
                  {sale.ecount_synced_at && ` · ${formatDate(sale.ecount_synced_at, 'yyyy.MM.dd HH:mm')}`}
                </p>
              </div>
              <Badge className="bg-blue-100 text-blue-700">연동 완료</Badge>
            </div>
          </Card>
        )}
      </div>
    </>
  );
}
