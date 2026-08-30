'use client';

import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { useBookInvoice, useCancelInvoice, useCancelOrder, useCompletePickup } from '@/hooks/use-orders';
import { InvoiceModal } from './invoice-modal';
import { OrderExchangeModal } from './order-exchange-modal';
import { AlertTriangle, Truck, ExternalLink, Store, RefreshCw } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import type { Order, OrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

interface Props {
  order: Order;
  items: OrderItem[];
}

export function OrderActionBar({ order, items }: Props) {
  const [invoiceOpen, setInvoiceOpen] = useState(false);
  const [showCancelOrder, setShowCancelOrder] = useState(false);
  const [showCancelInvoice, setShowCancelInvoice] = useState(false);
  const [showPushImweb, setShowPushImweb] = useState(false);
  const [showPickup, setShowPickup] = useState(false);
  const [showExchange, setShowExchange] = useState(false);
  const [checkingAlps, setCheckingAlps] = useState(false);
  const bookInvoice = useBookInvoice();
  const cancelInvoice = useCancelInvoice();
  const cancelOrder = useCancelOrder();
  const completePickup = useCompletePickup();

  const busy = bookInvoice.isPending || cancelInvoice.isPending || cancelOrder.isPending || completePickup.isPending || checkingAlps;

  // 주문 취소 확인 문구 — 상태별 경고 (집하 후 배송중은 강한 경고)
  const cancelOrderMessage = order.status === 'shipping'
    ? <>⚠️ <strong>이미 발송(집하)된 주문</strong>입니다. 취소하면 재고가 복구됩니다.<br />실제 물건이 배송 중이면 아임웹 반품·롯데 반송을 먼저 확인하세요.</>
    : order.invoice_number
      ? <>송장(<strong>{order.invoice_number}</strong>)이 발급된 주문입니다.<br />롯데 송장까지 정리하려면 <strong>[송장 취소]</strong>를 쓰세요. 그래도 주문을 취소하면 재고가 복구됩니다.<br />아임웹에서도 취소 처리하세요.</>
      : <>이 주문을 취소합니다. 재고가 복구됩니다.<br />아임웹에서도 취소 처리해주세요.</>;

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

  // 취소 상태 — 액션 없음 (배송완료는 '제품 교환' 위해 액션바 유지)
  if (order.status === 'cancelled') return null;

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
            <Button variant="secondary" size="sm" className="w-full" onClick={() => setShowPickup(true)} disabled={busy}>
              <Store size={14} />
              직접수령 완료
            </Button>
            <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => setShowCancelOrder(true)} disabled={busy}>
              주문 취소
            </Button>
          </>
        )}

        {/* 배송대기/배송중: 아임웹 연동 + 송장 취소 + 주문 취소 */}
        {(order.status === 'ready_to_ship' || order.status === 'shipping' || (order.status === 'pay_done' && order.invoice_number)) && (
          <>
            {/* 실제 롯데 송장이 있을 때만 — 아임웹 재연동 / 롯데 송장 취소 */}
            {order.invoice_number && (
              <>
                <Button variant="secondary" size="sm" className="w-full" onClick={() => setShowPushImweb(true)} disabled={busy}>
                  <ExternalLink size={14} />
                  아임웹 송장 연동
                </Button>
                <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => setShowCancelInvoice(true)} disabled={busy}>
                  송장 취소
                </Button>
              </>
            )}
            {/* 주문 취소 — 모든 진행 상태에서 '정리' 가능 (재고 복구). orphan(송장없음)도 이 버튼으로 정리 */}
            <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => setShowCancelOrder(true)} disabled={busy}>
              주문 취소
            </Button>
            <p className="text-[11px] text-neutral-400 px-1 leading-relaxed">
              {order.invoice_number
                ? '집하 전 취소는 [송장 취소]로 롯데 송장까지 정리 · 이미 발송됐거나 강제 정리는 [주문 취소]'
                : '롯데 송장이 없는 주문입니다. [주문 취소]로 정리하세요.'}
            </p>
          </>
        )}

        {/* 제품 교환 — 모든 진행/완료(취소 제외) 상태에서 가능. 매출·카드 불변, 상품/재고만 스왑 */}
        <Button variant="secondary" size="sm" className="w-full" onClick={() => setShowExchange(true)} disabled={busy}>
          <RefreshCw size={14} />
          제품 교환
        </Button>
        {order.exchanged_at && (
          <p className="text-[11px] text-emerald-600 px-1">✓ 교환 처리됨 — 아임웹 주문/결제는 그대로 유지됨</p>
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

      {/* 제품 교환 모달 */}
      {showExchange && (
        <OrderExchangeModal order={order} items={items} onClose={() => setShowExchange(false)} />
      )}

      {/* 주문 취소 확인 */}
      <ConfirmModal
        open={showCancelOrder}
        onClose={() => setShowCancelOrder(false)}
        onConfirm={async () => { await cancelOrder.mutateAsync(order.id); }}
        title="주문 취소"
        message={cancelOrderMessage}
        confirmLabel="주문 취소"
        variant="danger"
      />

      {/* 송장 취소 확인 */}
      <ConfirmModal
        open={showCancelInvoice}
        onClose={() => setShowCancelInvoice(false)}
        onConfirm={() => cancelInvoice.mutateAsync({ invNo: order.invoice_number!, orderId: order.id })}
        title="송장 취소"
        message={<>송장 <strong>{order.invoice_number}</strong>을 취소합니다.<br />ALPS 집하 전에만 가능합니다.</>}
        confirmLabel="송장 취소"
        variant="danger"
      />

      {/* 직접수령 완료 확인 */}
      <ConfirmModal
        open={showPickup}
        onClose={() => setShowPickup(false)}
        onConfirm={async () => { await completePickup.mutateAsync(order.id); }}
        title="직접수령 완료"
        message={<>이 주문을 <strong>직접수령(대면 픽업)</strong>으로 완료합니다.<br />송장 없이 배송완료로 마감됩니다. 아임웹에서도 수령 처리해주세요.</>}
        confirmLabel="직접수령 완료"
      />

      {/* 아임웹 송장 연동 확인 */}
      <ConfirmModal
        open={showPushImweb}
        onClose={() => setShowPushImweb(false)}
        onConfirm={handlePushImweb}
        title="아임웹 송장 연동"
        message={<>송장 <strong>{order.invoice_number}</strong>을 아임웹에 연동합니다.<br />아임웹에서 &quot;배송대기&quot; 상태여야 합니다.</>}
        confirmLabel="연동"
      />
    </>
  );
}
