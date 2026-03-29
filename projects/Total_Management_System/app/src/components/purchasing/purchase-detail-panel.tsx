'use client';

import { useState } from 'react';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { usePurchaseOrder, useUpdatePurchaseOrder } from '@/hooks/use-purchasing';
import { formatKRW, formatDate, calcVAT } from '@/lib/utils/format';
import { Truck } from 'lucide-react';

const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', ordered: '발주완료', deposit_paid: '선납완료',
  received: '입고완료', balance_paid: '잔금완료', cancelled: '취소',
};
const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600', ordered: 'bg-blue-100 text-blue-700',
  deposit_paid: 'bg-yellow-100 text-yellow-700', received: 'bg-green-100 text-green-700',
  balance_paid: 'bg-emerald-100 text-emerald-700', cancelled: 'bg-red-100 text-red-700',
};

interface Props {
  purchaseId: string;
}

export function PurchaseDetailPanel({ purchaseId }: Props) {
  const { data, isLoading } = usePurchaseOrder(purchaseId);
  const updatePO = useUpdatePurchaseOrder();
  const [depositInput, setDepositInput] = useState('');
  const [pendingAction, setPendingAction] = useState<{
    status: string; label: string; msg: string; variant?: 'danger' | 'default'; extra?: Record<string, unknown>;
  } | null>(null);

  if (isLoading) {
    return <div className="space-y-4"><Skeleton className="h-48" /><Skeleton className="h-32" /></div>;
  }
  if (!data?.order) {
    return (
      <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
        <Truck size={28} className="mb-2 opacity-40" />
        <p className="text-sm">발주 정보를 찾을 수 없습니다</p>
      </div>
    );
  }

  const { order: po, items } = data;
  const { supply, vat } = calcVAT(po.total_amount);

  async function handleAction(status: string, extra?: Record<string, unknown>) {
    await updatePO.mutateAsync({ id: purchaseId, status, ...extra });
  }

  return (
    <div className="flex-1 overflow-y-auto space-y-4 pr-1">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div>
          <h3 className="text-sm font-bold text-indigo-black">{po.po_number}</h3>
          <p className="text-xs text-neutral-500 mt-0.5">{po.supplier_name} · {formatDate(po.order_date)}</p>
        </div>
        <Badge className={STATUS_COLOR[po.status] || ''}>{STATUS_LABEL[po.status] || po.status}</Badge>
      </div>

      {/* 발주 정보 */}
      <Card>
        <div className="grid grid-cols-2 gap-3 text-sm">
          <div>
            <span className="text-xs text-neutral-500">매입처</span>
            <p className="font-semibold">{po.supplier_name}</p>
          </div>
          <div>
            <span className="text-xs text-neutral-500">발주일</span>
            <p>{formatDate(po.order_date)}</p>
          </div>
          {po.expected_date && (
            <div>
              <span className="text-xs text-neutral-500">입고 예정일</span>
              <p>{formatDate(po.expected_date)}</p>
            </div>
          )}
          {po.received_date && (
            <div>
              <span className="text-xs text-neutral-500">입고일</span>
              <p>{formatDate(po.received_date)}</p>
            </div>
          )}
        </div>
        {po.memo && <p className="mt-3 pt-3 border-t border-neutral-100 text-sm text-neutral-600">{po.memo}</p>}
      </Card>

      {/* 품목 */}
      <Card>
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">발주 품목</h4>
        <div className="space-y-2">
          {items.map((item) => (
            <div key={item.id} className="flex items-center justify-between py-1.5">
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
          <div className="flex justify-between text-sm font-bold">
            <span>합계</span>
            <span className="text-terracotta">{formatKRW(po.total_amount)}</span>
          </div>
          <div className="flex justify-between text-xs text-neutral-500">
            <span>공급가액 {formatKRW(supply)}</span>
            <span>부가세 {formatKRW(vat)}</span>
          </div>
        </div>
      </Card>

      {/* 결제 현황 */}
      <Card>
        <h4 className="text-xs font-semibold text-neutral-500 mb-2">결제 현황</h4>
        <div className="grid grid-cols-3 gap-3 text-center">
          <div>
            <p className="text-xs text-neutral-400">총액</p>
            <p className="text-sm font-bold">{formatKRW(po.total_amount)}</p>
          </div>
          <div>
            <p className="text-xs text-neutral-400">선납</p>
            <p className="text-sm font-bold text-blue-600">{formatKRW(po.deposit_amount)}</p>
          </div>
          <div>
            <p className="text-xs text-neutral-400">잔금</p>
            <p className={`text-sm font-bold ${po.balance_amount > 0 ? 'text-red-500' : 'text-green-600'}`}>
              {formatKRW(po.balance_amount)}
            </p>
          </div>
        </div>
      </Card>

      {/* 액션 */}
      {po.status !== 'cancelled' && po.status !== 'balance_paid' && (
        <Card>
          <h4 className="text-xs font-semibold text-neutral-500 mb-2">액션</h4>
          <div className="space-y-2">
            {po.status === 'draft' && (
              <Button className="w-full" onClick={() => setPendingAction({ status: 'ordered', label: '발주 확정', msg: '이 발주를 확정합니다.' })} disabled={updatePO.isPending}>
                발주 확정
              </Button>
            )}
            {(po.status === 'ordered' || po.status === 'draft') && (
              <div className="flex gap-2">
                <input type="number" value={depositInput} onChange={(e) => setDepositInput(e.target.value)}
                  placeholder={`선납금`}
                  className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40" />
                <Button variant="secondary" onClick={() => setPendingAction({
                  status: 'deposit_paid', label: '선납 처리',
                  msg: `선납금 ${formatKRW(parseInt(depositInput) || Math.round(po.total_amount / 2))}을 처리합니다.`,
                  extra: { deposit_amount: parseInt(depositInput) || Math.round(po.total_amount / 2) },
                })} disabled={updatePO.isPending}>
                  선납
                </Button>
              </div>
            )}
            {(po.status === 'ordered' || po.status === 'deposit_paid') && (
              <Button className="w-full bg-green-600 hover:bg-green-700" onClick={() => setPendingAction({
                status: 'received', label: '입고 확인',
                msg: `입고를 확인합니다. 재고가 자동으로 증가합니다.\n\n${items.map((i: { product_name: string; quantity: number }) => `• ${i.product_name} x${i.quantity}`).join('\n')}`,
              })} disabled={updatePO.isPending}>
                입고 확인 (재고 증가)
              </Button>
            )}
            {po.status === 'received' && po.balance_amount > 0 && (
              <Button className="w-full" onClick={() => setPendingAction({ status: 'balance_paid', label: '잔금 완료', msg: `잔금 ${formatKRW(po.balance_amount)}을 완료 처리합니다.` })} disabled={updatePO.isPending}>
                잔금 완료
              </Button>
            )}
            <Button variant="ghost" className="w-full text-red-500 hover:text-red-600" onClick={() => setPendingAction({ status: 'cancelled', label: '발주 취소', msg: '이 발주를 취소합니다.', variant: 'danger' })} disabled={updatePO.isPending}>
              발주 취소
            </Button>
          </div>
        </Card>
      )}

      {/* 확인 모달 */}
      {pendingAction && (
        <ConfirmModal
          open={!!pendingAction}
          onClose={() => setPendingAction(null)}
          onConfirm={() => handleAction(pendingAction.status, pendingAction.extra)}
          title={pendingAction.label}
          message={<span className="whitespace-pre-wrap">{pendingAction.msg}</span>}
          confirmLabel={pendingAction.label}
          variant={pendingAction.variant || 'default'}
        />
      )}
    </div>
  );
}
