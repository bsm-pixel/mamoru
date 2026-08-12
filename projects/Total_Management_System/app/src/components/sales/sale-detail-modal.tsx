'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useCancelSale, useUpdatePaymentStatus } from '@/hooks/use-sales';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { Hash, Ban, CheckCircle, AlertTriangle } from 'lucide-react';
import type { SaleChannel } from '@/lib/supabase/types';

const PAYMENT_METHOD_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
};

const PAYMENT_STATUS_COLOR: Record<string, string> = {
  paid: 'bg-green-100 text-green-700',
  unpaid: 'bg-red-100 text-red-700',
  partial: 'bg-yellow-100 text-yellow-700',
};

const PAYMENT_STATUS_LABEL: Record<string, string> = {
  paid: '결제완료', unpaid: '미결제', partial: '부분결제',
};

// 채널 칩 — 2026-07-17 4분류(매장/출장/톡/온라인) + 레거시 오프라인
const CHANNEL_CHIP: Record<string, { label: string; className: string }> = {
  store:   { label: '매장', className: 'bg-neutral-100 text-neutral-700' },
  field:   { label: '출장', className: 'bg-emerald-100 text-emerald-700' },
  talk:    { label: '톡',   className: 'bg-yellow-100 text-yellow-700' },
  online:  { label: '온라인(아임웹)', className: 'bg-blue-100 text-blue-700' },
  offline: { label: '오프라인', className: 'bg-neutral-100 text-neutral-500' }, // 레거시
};

interface Props {
  saleId: string | null;
  open: boolean;
  onClose: () => void;
}

export function SaleDetailModal({ saleId, open, onClose }: Props) {
  const { data, isLoading } = useSale(saleId || '');
  const cancelSale = useCancelSale();
  const updatePayment = useUpdatePaymentStatus();
  const [cancelMode, setCancelMode] = useState(false);
  const [cancelReason, setCancelReason] = useState('');

  if (!open || !saleId) return null;

  const handleCancel = async () => {
    await cancelSale.mutateAsync({ id: saleId, reason: cancelReason });
    setCancelMode(false);
    setCancelReason('');
    onClose();
  };

  const handleMarkPaid = () => {
    if (!data?.sale) return;
    updatePayment.mutate({
      id: saleId,
      payment_status: 'paid',
      // 실수납 = 소계 - 할인 (완납 시 할인 반영)
      paid_amount: Math.max(0, data.sale.total_amount - (data.sale.discount_amount || 0)),
    });
  };

  return (
    <Modal open={open} onClose={onClose} title="판매 상세" className="max-w-xl">
      {isLoading ? (
        <div className="space-y-3">
          <Skeleton className="h-20 w-full" />
          <Skeleton className="h-32 w-full" />
        </div>
      ) : !data ? (
        <p className="text-sm text-neutral-400 text-center py-8">판매 정보를 찾을 수 없습니다</p>
      ) : (
        <div className="space-y-4 max-h-[70vh] overflow-y-auto">
          <SaleInfo sale={data.sale} />
          <SaleItems items={data.items} serials={data.serials || []} />
          <SaleSummary sale={data.sale} />

          {/* 액션 영역 */}
          {data.sale.cancelled_at ? (
            <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-red-50 border border-red-100">
              <Ban size={14} className="text-red-500 shrink-0" />
              <div className="text-xs">
                <p className="font-semibold text-red-700">취소됨 — {formatDate(data.sale.cancelled_at)}</p>
                {data.sale.cancelled_reason && (
                  <p className="text-red-600 mt-0.5">{data.sale.cancelled_reason}</p>
                )}
              </div>
            </div>
          ) : (
            <div className="flex items-center gap-2 pt-2 border-t border-neutral-100">
              {/* 결제완료로 변경 */}
              {data.sale.payment_status !== 'paid' && (
                <Button
                  size="sm"
                  onClick={handleMarkPaid}
                  disabled={updatePayment.isPending}
                >
                  <CheckCircle size={14} />
                  {updatePayment.isPending ? '처리 중...' : '결제완료로 변경'}
                </Button>
              )}

              {/* 판매 취소 */}
              {!cancelMode ? (
                <Button
                  variant="ghost"
                  size="sm"
                  className="text-red-600 hover:bg-red-50"
                  onClick={() => setCancelMode(true)}
                >
                  <Ban size={14} />
                  판매 취소
                </Button>
              ) : (
                <div className="flex-1 space-y-2">
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
          )}
        </div>
      )}
    </Modal>
  );
}

/* --- 서브 컴포넌트 --- */

function SaleInfo({ sale }: { sale: Record<string, unknown> }) {
  const s = sale as unknown as import('@/lib/supabase/types').OfflineSale;
  const channel = CHANNEL_CHIP[(s.sale_channel || 'offline') as SaleChannel] || CHANNEL_CHIP.offline;

  return (
    <div>
      <div className="flex items-center gap-2 mb-2">
        <span className="text-sm font-bold text-indigo-black">{s.sale_number}</span>
        <Badge className={PAYMENT_STATUS_COLOR[s.payment_status] || ''}>
          {PAYMENT_STATUS_LABEL[s.payment_status] || s.payment_status}
        </Badge>
        <Badge className={channel.className}>{channel.label}</Badge>
      </div>
      <div className="grid grid-cols-2 gap-2 text-sm">
        <div>
          <span className="text-xs text-neutral-500">고객명</span>
          <p className="font-semibold">{s.customer_name}</p>
        </div>
        <div>
          <span className="text-xs text-neutral-500">연락처</span>
          <p>{formatPhone(s.customer_phone) || '-'}</p>
        </div>
        <div>
          <span className="text-xs text-neutral-500">판매일</span>
          <p>{formatDate(s.sale_date, 'yyyy.MM.dd')}</p>
        </div>
        <div>
          <span className="text-xs text-neutral-500">결제방법</span>
          <p>{PAYMENT_METHOD_LABEL[s.payment_method] || s.payment_method}</p>
        </div>
      </div>
      {s.memo && (
        <p className="mt-2 pt-2 border-t border-neutral-100 text-sm text-neutral-600">{s.memo}</p>
      )}
    </div>
  );
}

function SaleItems({ items, serials }: {
  items: import('@/lib/supabase/types').OfflineSaleItem[];
  serials: Array<{ id: string; serial_number: string; product_id: string | null; sale_item_id?: string | null }>;
}) {
  return (
    <div className="border-t border-neutral-100 pt-3">
      <h4 className="text-xs font-semibold text-neutral-500 mb-2">판매 항목</h4>
      <div className="space-y-2">
        {items.map((item) => {
          // 시리얼 매칭: 1순위 sale_item_id 정확 / 2순위 product_id fallback (레거시 데이터용)
          // (2026-05-17 fix — 다중 상품 판매 시리얼 잘못 표시 버그)
          const bySaleItem = serials.filter((s) => s.sale_item_id === item.id);
          const itemSerials = bySaleItem.length > 0
            ? bySaleItem
            : serials.filter((s) => !s.sale_item_id && s.product_id === item.product_id);
          return (
            <div key={item.id} className="py-1.5">
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
    </div>
  );
}

function SaleSummary({ sale }: { sale: Record<string, unknown> }) {
  const s = sale as unknown as import('@/lib/supabase/types').OfflineSale;

  return (
    <div className="border-t border-neutral-200 pt-3 space-y-1">
      <div className="flex justify-between text-sm">
        <span className="text-neutral-500">소계</span>
        <span>{formatKRW(s.total_amount)}</span>
      </div>
      {s.discount_amount > 0 && (
        <div className="flex justify-between text-sm">
          <span className="text-neutral-500">할인</span>
          <span className="text-red-600">-{formatKRW(s.discount_amount)}</span>
        </div>
      )}
      <div className="flex justify-between text-sm font-bold">
        <span>결제 금액</span>
        <span className="text-stone-900">{formatKRW(s.paid_amount)}</span>
      </div>
      {s.supply_amount > 0 && (
        <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
          <div className="flex justify-between text-xs text-neutral-500">
            <span>공급가액</span>
            <span>{formatKRW(s.supply_amount)}</span>
          </div>
          <div className="flex justify-between text-xs text-neutral-500">
            <span>부가세 (10%)</span>
            <span>{formatKRW(s.vat_amount)}</span>
          </div>
        </div>
      )}
    </div>
  );
}
