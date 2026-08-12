'use client';

import { use, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useCancelSale, useUpdatePaymentStatus } from '@/hooks/use-sales';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { ArrowLeft, Hash, CheckCircle, Ban, AlertTriangle } from 'lucide-react';
import type { SaleChannel } from '@/lib/supabase/types';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드',
  cash: '현금',
  transfer: '계좌이체',
  mixed: '복합',
};

const CHANNEL_CHIP: Record<string, { label: string; className: string }> = {
  offline: { label: '오프라인', className: 'bg-neutral-100 text-neutral-600' },
  online:  { label: '온라인',  className: 'bg-blue-100 text-blue-700' },
  talk:    { label: '온라인상담',  className: 'bg-yellow-100 text-yellow-700' },
};

export default function SaleDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();
  const { data, isLoading } = useSale(id);
  const cancelSale = useCancelSale();
  const updatePayment = useUpdatePaymentStatus();
  const [cancelMode, setCancelMode] = useState(false);
  const [cancelReason, setCancelReason] = useState('');

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

  const { sale, items, serials = [] } = data;
  const isCancelled = !!sale.cancelled_at;
  const channel = CHANNEL_CHIP[(sale.sale_channel || 'offline') as SaleChannel] || CHANNEL_CHIP.offline;

  const handleCancel = async () => {
    await cancelSale.mutateAsync({ id, reason: cancelReason });
    setCancelMode(false);
    setCancelReason('');
  };

  const handleMarkPaid = () => {
    // 실수납 = 소계 - 할인 (완납 시 할인 반영)
    updatePayment.mutate({ id, payment_status: 'paid', paid_amount: Math.max(0, sale.total_amount - (sale.discount_amount || 0)) });
  };

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
            <div className="flex items-center gap-2">
              <h3 className="text-sm font-bold text-stone-900">{sale.sale_number}</h3>
              <Badge className={channel.className}>{channel.label}</Badge>
            </div>
            {isCancelled ? (
              <Badge className="bg-neutral-200 text-neutral-500">취소</Badge>
            ) : (
              <Badge className={
                sale.payment_status === 'paid' ? 'bg-green-100 text-green-700'
                  : sale.payment_status === 'unpaid' ? 'bg-red-100 text-red-700'
                  : 'bg-yellow-100 text-yellow-700'
              }>
                {sale.payment_status === 'paid' ? '결제완료' : sale.payment_status === 'unpaid' ? '미결제' : '부분결제'}
              </Badge>
            )}
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
          <h3 className="text-sm font-bold text-stone-900 mb-3">판매 항목</h3>
          <div className="space-y-2">
            {items.map((item) => {
              // 시리얼 매칭: 1순위 sale_item_id 정확 매칭 / 2순위 product_id fallback (레거시 데이터용)
              // (2026-05-17 fix — 다중 상품 판매에서 시리얼 잘못 표시되던 버그)
              const bySaleItem = serials.filter((s) => s.sale_item_id === item.id);
              const itemSerials = bySaleItem.length > 0
                ? bySaleItem
                : serials.filter((s) => !s.sale_item_id && s.product_id === item.product_id);
              return (
                <div key={item.id} className="py-2 border-b border-neutral-50 last:border-0">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm font-medium">{item.product_name}</p>
                      <p className="text-xs text-neutral-500">
                        {item.sku && `${item.sku} · `}{formatKRW(item.unit_price)} x {item.quantity}
                      </p>
                    </div>
                    <span className="text-sm font-bold">{formatKRW(item.total_price)}</span>
                  </div>
                  {itemSerials.length > 0 && (
                    <div className="mt-1 flex flex-wrap gap-1">
                      {itemSerials.map((s) => (
                        <span key={s.id} className="inline-flex items-center gap-0.5 text-[10px] font-mono bg-neutral-100 text-neutral-600 px-1.5 py-0.5 rounded">
                          <Hash size={8} />{s.serial_number}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
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
              <span className="text-stone-900">{formatKRW(sale.paid_amount)}</span>
            </div>
            {/* VAT 분리 표시 */}
            {sale.supply_amount > 0 && (
              <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
                <div className="flex justify-between text-xs text-neutral-500">
                  <span>공급가액</span>
                  <span>{formatKRW(sale.supply_amount)}</span>
                </div>
                <div className="flex justify-between text-xs text-neutral-500">
                  <span>부가세 (10%)</span>
                  <span>{formatKRW(sale.vat_amount)}</span>
                </div>
              </div>
            )}
          </div>
        </Card>

        {/* 취소 정보 또는 액션 버튼 */}
        {isCancelled ? (
          <Card>
            <div className="flex items-center gap-2">
              <Ban size={14} className="text-red-500 shrink-0" />
              <div className="text-sm">
                <p className="font-semibold text-red-700">취소됨 — {formatDate(sale.cancelled_at!)}</p>
                {sale.cancelled_reason && (
                  <p className="text-red-600 mt-0.5">{sale.cancelled_reason}</p>
                )}
              </div>
            </div>
          </Card>
        ) : (
          <Card>
            <div className="flex items-center gap-2 flex-wrap">
              {sale.payment_status !== 'paid' && (
                <Button size="sm" onClick={handleMarkPaid} disabled={updatePayment.isPending}>
                  <CheckCircle size={14} />
                  {updatePayment.isPending ? '처리 중...' : '결제완료로 변경'}
                </Button>
              )}
              {!cancelMode ? (
                <Button variant="ghost" size="sm" className="text-red-600 hover:bg-red-50" onClick={() => setCancelMode(true)}>
                  <Ban size={14} />
                  판매 취소
                </Button>
              ) : (
                <div className="w-full space-y-2 mt-2">
                  <div className="flex items-center gap-1.5 text-xs text-red-600">
                    <AlertTriangle size={12} />
                    <span>시리얼/재고가 복원됩니다. 취소 사유를 입력해주세요.</span>
                  </div>
                  <input
                    type="text"
                    value={cancelReason}
                    onChange={(e) => setCancelReason(e.target.value)}
                    placeholder="취소 사유 (선택)"
                    className="w-full h-8 px-3 rounded-lg border border-red-200 bg-red-50/50 text-sm focus:outline-none focus:ring-2 focus:ring-red-300"
                  />
                  <div className="flex gap-2">
                    <Button
                      size="sm"
                      className="bg-red-600 hover:bg-red-700 text-white"
                      onClick={handleCancel}
                      disabled={cancelSale.isPending}
                    >
                      {cancelSale.isPending ? '취소 처리 중...' : '취소 확정'}
                    </Button>
                    <Button variant="ghost" size="sm" onClick={() => { setCancelMode(false); setCancelReason(''); }}>
                      돌아가기
                    </Button>
                  </div>
                </div>
              )}
            </div>
          </Card>
        )}

      </div>
    </>
  );
}
