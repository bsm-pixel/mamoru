'use client';

import { useState } from 'react';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSale, useCancelSale, useUpdatePaymentStatus, useUpdateSaleMemo, useEditSale } from '@/hooks/use-sales';
import { formatKRW, formatDate, formatPhone } from '@/lib/utils/format';
import { Hash, Ban, CheckCircle, AlertTriangle, Pencil, Save, FileText } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
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
  const editSale = useEditSale();
  const [cancelMode, setCancelMode] = useState(false);
  const [cancelReason, setCancelReason] = useState('');
  const [showPaidConfirm, setShowPaidConfirm] = useState(false);
  const [showCancelConfirm, setShowCancelConfirm] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
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
          {(() => {
            // 시리얼-항목 매칭: product_id 있으면 매칭, 없으면 순서대로 배분
            type Sr = { id: string; serial_number: string; product_id: string | null };
            const allSerials = serials as Sr[];
            const matched = allSerials.filter((sr) => sr.product_id);
            const unmatched = allSerials.filter((sr) => !sr.product_id);
            let unmatchedIdx = 0;

            return (items as OfflineSaleItem[]).map((item) => {
              // product_id 매칭
              const byProduct = matched.filter((sr) => sr.product_id === item.product_id);
              // NULL 시리얼은 항목 수량만큼 순서대로 배분
              const byOrder: Sr[] = [];
              if (byProduct.length === 0) {
                const take = Math.min(item.quantity, unmatched.length - unmatchedIdx);
                for (let i = 0; i < take; i++) {
                  byOrder.push(unmatched[unmatchedIdx++]);
                }
              }
              const itemSerials = [...byProduct, ...byOrder];

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
                      {itemSerials.map((sr) => (
                        <span key={sr.id} className="inline-flex items-center gap-0.5 text-[10px] font-mono bg-neutral-100 text-neutral-600 px-1.5 py-0.5 rounded">
                          <Hash size={8} />{sr.serial_number}
                        </span>
                      ))}
                    </div>
                  )}
                </div>
              );
            });
          })()}
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

      {/* 거래명세서 + 수정 */}
      <div className="flex gap-2">
        <button
          onClick={() => window.open(`/reports/transaction?sale_id=${saleId}`, '_blank')}
          className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
        >
          <FileText size={14} />
          거래명세서
        </button>
        {!s.cancelled_at && (
          <button
            onClick={() => setShowEditModal(true)}
            className="flex-1 flex items-center justify-center gap-2 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600 hover:bg-neutral-50 transition"
          >
            <Pencil size={14} />
            수정
          </button>
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
            <Button size="sm" onClick={() => setShowPaidConfirm(true)} disabled={updatePayment.isPending}>
              <CheckCircle size={14} />
              {updatePayment.isPending ? '처리 중...' : '결제완료로 변경'}
            </Button>
          )}
          <button onClick={() => setShowCancelConfirm(true)} className="text-xs text-red-500 hover:text-red-700 transition">
            판매 취소
          </button>
        </div>
      )}

      {/* 결제완료 확인 모달 */}
      <ConfirmModal
        open={showPaidConfirm}
        onClose={() => setShowPaidConfirm(false)}
        onConfirm={() => updatePayment.mutateAsync({ id: saleId, payment_status: 'paid', paid_amount: s.total_amount })}
        title="결제완료 처리"
        message={<>{s.customer_name}님의 결제를 완료 처리합니다.<br />금액: {formatKRW(s.total_amount)}</>}
        confirmLabel="결제완료"
      />

      {/* 판매 취소 확인 모달 */}
      <ConfirmModal
        open={showCancelConfirm}
        onClose={() => { setShowCancelConfirm(false); setCancelReason(''); }}
        onConfirm={() => cancelSale.mutateAsync({ id: saleId, reason: cancelReason })}
        title="판매 취소"
        message={
          <div className="space-y-2">
            <div className="flex items-center gap-1.5 text-xs text-red-600">
              <AlertTriangle size={12} />
              <span>시리얼/재고가 복원됩니다. 이 작업은 되돌릴 수 없습니다.</span>
            </div>
            <input
              type="text"
              value={cancelReason}
              onChange={(e) => setCancelReason(e.target.value)}
              placeholder="취소 사유 (선택)"
              className="w-full h-8 px-3 rounded-lg border border-red-200 bg-red-50/50 text-sm focus:outline-none focus:ring-2 focus:ring-red-300"
            />
          </div>
        }
        confirmLabel="취소 확정"
        variant="danger"
      />

      {/* 판매 수정 모달 */}
      {showEditModal && (
        <EditSaleModal
          sale={s}
          onClose={() => setShowEditModal(false)}
          onSave={(fields) => editSale.mutateAsync({ id: saleId, ...fields }).then(() => setShowEditModal(false))}
          isPending={editSale.isPending}
        />
      )}
    </div>
  );
}

/** 판매 수정 모달 */
function EditSaleModal({ sale, onClose, onSave, isPending }: {
  sale: OfflineSale;
  onClose: () => void;
  onSave: (fields: Record<string, unknown>) => Promise<unknown>;
  isPending: boolean;
}) {
  const [totalAmount, setTotalAmount] = useState(sale.total_amount);
  const [discountAmount, setDiscountAmount] = useState(sale.discount_amount || 0);
  const [paymentMethod, setPaymentMethod] = useState<string>(sale.payment_method);
  const [saleDate, setSaleDate] = useState(sale.sale_date);

  const METHODS = [
    { value: 'card', label: '카드' },
    { value: 'cash', label: '현금' },
    { value: 'transfer', label: '이체' },
    { value: 'mixed', label: '복합' },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-[90vw] max-w-[400px] p-5 space-y-4" onClick={(e) => e.stopPropagation()}>
        <h3 className="text-sm font-bold text-neutral-800">판매 정보 수정</h3>

        <div>
          <label className="text-xs text-neutral-500 mb-1 block">판매일</label>
          <input type="date" value={saleDate} max={new Date().toISOString().slice(0, 10)}
            onChange={(e) => setSaleDate(e.target.value)}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>

        <div>
          <label className="text-xs text-neutral-500 mb-1 block">총 금액</label>
          <input type="number" value={totalAmount}
            onChange={(e) => setTotalAmount(Number(e.target.value))}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>

        <div>
          <label className="text-xs text-neutral-500 mb-1 block">할인</label>
          <input type="number" value={discountAmount}
            onChange={(e) => setDiscountAmount(Number(e.target.value))}
            className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>

        <div>
          <label className="text-xs text-neutral-500 mb-1 block">결제방법</label>
          <div className="flex gap-2">
            {METHODS.map((m) => (
              <button key={m.value}
                onClick={() => setPaymentMethod(m.value)}
                className={`flex-1 py-1.5 text-xs rounded-md border transition ${
                  paymentMethod === m.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                }`}
              >{m.label}</button>
            ))}
          </div>
        </div>

        <div className="flex gap-2 pt-2">
          <button onClick={onClose} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">
            취소
          </button>
          <button
            onClick={() => onSave({
              total_amount: totalAmount,
              discount_amount: discountAmount,
              payment_method: paymentMethod,
              sale_date: saleDate,
            })}
            disabled={isPending}
            className="flex-1 py-2 rounded-lg bg-neutral-900 text-white text-sm font-medium disabled:opacity-50"
          >
            {isPending ? '저장 중...' : '저장'}
          </button>
        </div>
      </div>
    </div>
  );
}
