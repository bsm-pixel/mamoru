'use client';

import { useState } from 'react';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useCancelSale, useUpdatePaymentStatus, useUpdateSaleMemo } from '@/hooks/use-sales';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { Hash, Ban, CheckCircle, AlertTriangle, Pencil, Save } from 'lucide-react';
import type { SaleChannel, OfflineSale, OfflineSaleItem } from '@/lib/supabase/types';

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

const CHANNEL_CHIP: Record<string, { label: string; className: string }> = {
  offline: { label: '오프라인', className: 'bg-neutral-100 text-neutral-600' },
  online:  { label: '온라인',  className: 'bg-blue-100 text-blue-700' },
  talk:    { label: '톡상담',  className: 'bg-yellow-100 text-yellow-700' },
};

interface Props {
  saleId: string;
}

export function SaleDetailPanel({ saleId }: Props) {
  const { data, isLoading } = useSale(saleId);
  const cancelSale = useCancelSale();
  const updatePayment = useUpdatePaymentStatus();
  const updateMemo = useUpdateSaleMemo();
  const [cancelMode, setCancelMode] = useState(false);
  const [cancelReason, setCancelReason] = useState('');
  const [editingMemo, setEditingMemo] = useState(false);
  const [memoValue, setMemoValue] = useState('');

  if (isLoading) {
    return (
      <div className="p-4 space-y-3">
        <Skeleton className="h-20 w-full" />
        <Skeleton className="h-32 w-full" />
      </div>
    );
  }

  if (!data) {
    return <p className="text-sm text-neutral-400 text-center py-8">판매 정보를 찾을 수 없습니다</p>;
  }

  const { sale, items, serials = [] } = data;
  const s = sale as unknown as OfflineSale;
  const channel = CHANNEL_CHIP[(s.sale_channel || 'offline') as SaleChannel] || CHANNEL_CHIP.offline;

  const handleCancel = async () => {
    await cancelSale.mutateAsync({ id: saleId, reason: cancelReason });
    setCancelMode(false);
    setCancelReason('');
  };

  const handleMarkPaid = () => {
    updatePayment.mutate({
      id: saleId,
      payment_status: 'paid',
      paid_amount: s.total_amount,
    });
  };

  const handleSaveMemo = () => {
    updateMemo.mutate({ id: saleId, memo: memoValue });
    setEditingMemo(false);
  };

  return (
    <div className="p-4 space-y-4">
      {/* 헤더 */}
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
            {/* 복합 결제 상세 */}
            {s.payment_method === 'mixed' && (() => {
              // eslint-disable-next-line @typescript-eslint/no-explicit-any
              const pd = (s as any).payment_detail as Record<string, number> | null;
              if (!pd) return null;
              return (
                <div className="mt-1 text-xs text-neutral-500 space-y-0.5">
                  {pd.card > 0 && <p>카드 {formatKRW(pd.card)}</p>}
                  {pd.cash > 0 && <p>현금 {formatKRW(pd.cash)}</p>}
                  {pd.transfer > 0 && <p>이체 {formatKRW(pd.transfer)}</p>}
                </div>
              );
            })()}
          </div>
        </div>
        {/* 메모 인라인 편집 */}
        <div className="mt-2 pt-2 border-t border-neutral-100">
          {editingMemo ? (
            <div className="flex gap-2 items-center">
              <input
                type="text"
                value={memoValue}
                onChange={(e) => setMemoValue(e.target.value)}
                placeholder="메모 입력"
                className="flex-1 h-8 px-2 rounded border border-neutral-200 bg-warm-ivory text-sm focus:outline-none focus:ring-2 focus:ring-terracotta/40"
                autoFocus
              />
              <button onClick={handleSaveMemo} className="p-1 hover:bg-neutral-100 rounded"><Save size={14} className="text-terracotta" /></button>
              <button onClick={() => setEditingMemo(false)} className="p-1 hover:bg-neutral-100 rounded text-xs text-neutral-500">취소</button>
            </div>
          ) : (
            <div className="flex items-center gap-1.5 text-sm text-neutral-600">
              <span>{s.memo || '메모 없음'}</span>
              <button onClick={() => { setMemoValue(s.memo || ''); setEditingMemo(true); }} className="p-0.5 hover:bg-neutral-100 rounded">
                <Pencil size={11} className="text-neutral-400" />
              </button>
            </div>
          )}
        </div>
      </div>

      {/* 판매 항목 */}
      <div className="border-t border-neutral-100 pt-3">
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">판매 항목</h4>
        <div className="space-y-2">
          {(items as OfflineSaleItem[]).map((item) => {
            const itemSerials = serials.filter((sr: { product_id: string }) => sr.product_id === item.product_id);
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
                    {itemSerials.map((sr: { id: string; serial_number: string }) => (
                      <span key={sr.id} className="inline-flex items-center gap-0.5 text-[10px] font-mono bg-neutral-100 text-neutral-600 px-1.5 py-0.5 rounded">
                        <Hash size={8} />{sr.serial_number}
                      </span>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {/* 합계 */}
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
          <span className="text-terracotta">{formatKRW(s.paid_amount)}</span>
        </div>
        {s.payment_status !== 'paid' && (
          <div className="flex justify-between text-sm font-semibold text-red-500">
            <span>미수금</span>
            <span>{formatKRW(s.total_amount - s.paid_amount)}</span>
          </div>
        )}
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

      {/* 액션 */}
      {s.cancelled_at ? (
        <div className="flex items-center gap-2 px-3 py-2.5 rounded-lg bg-red-50 border border-red-100">
          <Ban size={14} className="text-red-500 shrink-0" />
          <div className="text-xs">
            <p className="font-semibold text-red-700">취소됨 — {formatDate(s.cancelled_at)}</p>
            {s.cancelled_reason && <p className="text-red-600 mt-0.5">{s.cancelled_reason}</p>}
          </div>
        </div>
      ) : (
        <div className="flex items-center gap-2 pt-2 border-t border-neutral-100">
          {s.payment_status !== 'paid' && (
            <Button size="sm" onClick={handleMarkPaid} disabled={updatePayment.isPending}>
              <CheckCircle size={14} />
              {updatePayment.isPending ? '처리 중...' : '결제완료로 변경'}
            </Button>
          )}
          {!cancelMode ? (
            <button onClick={() => setCancelMode(true)} className="text-xs text-red-500 hover:text-red-700 transition">
              판매 취소
            </button>
          ) : (
            <div className="flex-1 space-y-2">
              <div className="flex items-center gap-1.5 text-xs text-red-600">
                <AlertTriangle size={12} />
                <span>시리얼/재고가 복원됩니다.</span>
              </div>
              <input
                type="text"
                value={cancelReason}
                onChange={(e) => setCancelReason(e.target.value)}
                placeholder="취소 사유 (선택)"
                className="w-full h-8 px-3 rounded-lg border border-red-200 bg-red-50/50 text-sm focus:outline-none focus:ring-2 focus:ring-red-300"
              />
              <div className="flex gap-2">
                <Button size="sm" className="bg-red-600 hover:bg-red-700 text-white" onClick={handleCancel} disabled={cancelSale.isPending}>
                  {cancelSale.isPending ? '취소 처리 중...' : '취소 확정'}
                </Button>
                <Button variant="ghost" size="sm" onClick={() => { setCancelMode(false); setCancelReason(''); }}>돌아가기</Button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
