'use client';

import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { useBookInvoice, useCancelInvoice, useCancelOrder } from '@/hooks/use-orders';
import { InvoiceModal } from './invoice-modal';
import { AlertTriangle, Truck, ExternalLink } from 'lucide-react';
import type { Order, OrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

interface Props {
  order: Order;
  items: OrderItem[];
}

export function OrderActionBar({ order, items }: Props) {
  const [invoiceOpen, setInvoiceOpen] = useState(false);
  const [cancelConfirm, setCancelConfirm] = useState(false);
  const [checkingAlps, setCheckingAlps] = useState(false);
  const bookInvoice = useBookInvoice();
  const cancelInvoice = useCancelInvoice();
  const cancelOrder = useCancelOrder();

  const busy = bookInvoice.isPending || cancelInvoice.isPending || cancelOrder.isPending || checkingAlps;

  async function handleCheckAlpsCancel() {
    if (!order.invoice_number) return;
    setCheckingAlps(true);
    try {
      const res = await fetch('/api/lotte/check-cancel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ orderId: order.id, invNo: order.invoice_number }),
      });
      const data = await res.json();
      if (data.cancelled) {
        toast.success('취소 확인 완료');
      } else {
        toast('아직 ALPS에서 취소되지 않았습니다', { icon: '⏳' });
      }
    } catch {
      toast.error('ALPS 확인 실패');
    } finally {
      setCheckingAlps(false);
    }
  }

  async function handlePushImweb() {
    try {
      const res = await fetch('/api/imweb/push-invoice', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ orderId: order.id }),
      });
      const data = await res.json();
      if (data.success) {
        toast.success('아임웹 송장 연동 완료');
      } else if (data.needsManual) {
        toast('아임웹에서 "배송대기 처리" 먼저 진행해주세요', { icon: '⚠️', duration: 5000 });
      } else {
        toast.error(data.error || '연동 실패');
      }
    } catch {
      toast.error('아임웹 연동 실패');
    }
  }

  // 완료/취소 상태 — 액션 없음
  if (['delivered', 'cancelled'].includes(order.status)) return null;

  return (
    <>
      <div className="space-y-2 pt-3 border-t border-neutral-100">
        {/* cancel_pending: 경고 + 확인 버튼 */}
        {order.status === 'cancel_pending' && (
          <>
            <div className="flex items-start gap-2 p-3 rounded-lg bg-yellow-50 border border-yellow-200">
              <AlertTriangle size={16} className="text-yellow-600 shrink-0 mt-0.5" />
              <div className="text-xs text-yellow-700">
                <p className="font-semibold">ALPS 집하취소 필요</p>
                <p className="mt-0.5">송장 {order.invoice_number}을 ALPS에서 직접 취소해주세요</p>
              </div>
            </div>
            <Button size="sm" className="w-full" onClick={handleCheckAlpsCancel} disabled={busy} loading={checkingAlps}>
              ALPS 취소 확인
            </Button>
          </>
        )}

        {/* pay_done 송장 없음: 송장 생성 + 주문 취소 */}
        {order.status === 'pay_done' && !order.invoice_number && (
          <>
            <Button size="sm" className="w-full" onClick={() => setInvoiceOpen(true)} disabled={busy}>
              <Truck size={14} />
              송장 생성
            </Button>
            {!cancelConfirm ? (
              <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => setCancelConfirm(true)} disabled={busy}>
                주문 취소
              </Button>
            ) : (
              <div className="flex gap-2">
                <Button variant="ghost" size="sm" className="flex-1" onClick={() => setCancelConfirm(false)}>
                  아니오
                </Button>
                <Button variant="danger" size="sm" className="flex-1" onClick={() => { cancelOrder.mutate(order.id); setCancelConfirm(false); }} disabled={busy}>
                  취소 확인
                </Button>
              </div>
            )}
          </>
        )}

        {/* shipping: 아임웹 연동 + 송장 취소 */}
        {(order.status === 'shipping' || (order.status === 'pay_done' && order.invoice_number)) && (
          <>
            <Button variant="secondary" size="sm" className="w-full" onClick={handlePushImweb} disabled={busy}>
              <ExternalLink size={14} />
              아임웹 송장 연동
            </Button>
            <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => cancelInvoice.mutate({ invNo: order.invoice_number!, orderId: order.id })} disabled={busy}>
              송장 취소
            </Button>
          </>
        )}
      </div>

      {/* 송장 생성 모달 */}
      {invoiceOpen && (
        <InvoiceModal
          open={invoiceOpen}
          onClose={() => setInvoiceOpen(false)}
          order={order}
          items={items}
        />
      )}
    </>
  );
}
