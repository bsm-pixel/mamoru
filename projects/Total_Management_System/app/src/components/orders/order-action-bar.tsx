'use client';

import { useState, type ReactNode } from 'react';
import { Button } from '@/components/ui/button';
import { useBookInvoice, useCancelInvoice, useCancelOrder, useCompletePickup } from '@/hooks/use-orders';
import { InvoiceModal } from './invoice-modal';
import { OrderExchangeModal } from './order-exchange-modal';
import { AlertTriangle, Truck, ExternalLink, Store, RefreshCw } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import type { Order, OrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

export interface OrderActions {
  /** 상단 「다음 할 일」에 넣을 주 액션이 있는지 */
  hasPrimary: boolean;
  /** 상단 주 액션(긍정 진행): 송장 생성 · 직접수령 · 아임웹 연동 · ALPS 취소 확인 */
  primary: ReactNode;
  /** 하단 부차/파괴 액션: 송장 취소 · 주문 취소 · 제품 교환 */
  secondary: ReactNode;
  /** 모달 모음 — 패널에 딱 1번만 렌더 */
  modals: ReactNode;
}

/**
 * 주문 상세 액션을 상단(주)·하단(부차)·모달로 분리해 돌려주는 훅.
 * 상태(모달 open)를 한 곳에서만 소유 → 버튼은 상/하단 2곳에 두되 모달은 1벌만 렌더(중복 방지).
 */
export function useOrderActions(order: Order | null | undefined, items: OrderItem[]): OrderActions {
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

  // 로딩 등으로 아직 주문이 없으면 빈 액션 (rules-of-hooks: 훅은 위에서 이미 모두 호출됨)
  if (!order) return { hasPrimary: false, primary: null, secondary: null, modals: null };
  const ord = order; // non-null 별칭 — 아래 클로저(핸들러)에서 narrowing 유지용

  // 주문 취소 확인 문구 — 상태별 경고 (집하 후 배송중은 강한 경고)
  const cancelOrderMessage = order.status === 'shipping'
    ? <>⚠️ <strong>이미 발송(집하)된 주문</strong>입니다. 취소하면 재고가 복구됩니다.<br />실제 물건이 배송 중이면 아임웹 반품·롯데 반송을 먼저 확인하세요.</>
    : order.invoice_number
      ? <>송장(<strong>{order.invoice_number}</strong>)이 발급된 주문입니다.<br />롯데 송장까지 정리하려면 <strong>[송장 취소]</strong>를 쓰세요. 그래도 주문을 취소하면 재고가 복구됩니다.<br />아임웹에서도 취소 처리하세요.</>
      : <>이 주문을 취소합니다. 재고가 복구됩니다.<br />아임웹에서도 취소 처리해주세요.</>;

  async function handleCheckAlpsCancel() {
    if (!ord.invoice_number) return;
    setCheckingAlps(true);
    try {
      const res = await fetch('/api/lotte/check-cancel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ orderId: ord.id, invNo: ord.invoice_number }),
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
        body: JSON.stringify({ orderId: ord.id }),
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

  // 취소 상태 — 액션 없음
  if (order.status === 'cancelled') {
    return { hasPrimary: false, primary: null, secondary: null, modals: null };
  }

  // 상태 판정 (상호배타)
  const showInvoiceCreate = order.status === 'pay_done' && !order.invoice_number;
  const canPushImweb = !!order.invoice_number &&
    (order.status === 'ready_to_ship' || order.status === 'shipping' || order.status === 'pay_done');
  const showAlpsCheck = order.status === 'cancel_pending';
  const hasPrimary = showInvoiceCreate || canPushImweb || showAlpsCheck;

  // ── 상단: 주 액션(긍정 진행) ──
  const primary = (
    <>
      {showAlpsCheck && (
        <>
          <div className="flex items-start gap-2 p-2.5 rounded-lg bg-yellow-50 border border-yellow-200">
            <AlertTriangle size={15} className="text-yellow-600 shrink-0 mt-0.5" />
            <div className="text-[11px] text-yellow-700">
              <p className="font-semibold">ALPS 집하취소 필요</p>
              <p className="mt-0.5">송장 {order.invoice_number}을 ALPS에서 직접 취소해주세요</p>
            </div>
          </div>
          <Button size="sm" className="w-full" onClick={handleCheckAlpsCancel} disabled={busy} loading={checkingAlps}>
            ALPS 취소 확인
          </Button>
        </>
      )}
      {showInvoiceCreate && (
        <>
          <Button size="sm" className="w-full" onClick={() => setInvoiceOpen(true)} disabled={busy}>
            <Truck size={14} />
            송장 생성
          </Button>
          <Button variant="secondary" size="sm" className="w-full" onClick={() => setShowPickup(true)} disabled={busy}>
            <Store size={14} />
            직접수령 완료
          </Button>
        </>
      )}
      {canPushImweb && (
        <Button variant="secondary" size="sm" className="w-full" onClick={() => setShowPushImweb(true)} disabled={busy}>
          <ExternalLink size={14} />
          아임웹 송장 연동
        </Button>
      )}
    </>
  );

  // ── 하단: 부차/파괴 액션 ──
  const secondary = (
    <div className="space-y-2 pt-3 border-t border-neutral-100">
      {showInvoiceCreate && (
        <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => setShowCancelOrder(true)} disabled={busy}>
          주문 취소
        </Button>
      )}
      {canPushImweb && (
        <>
          {order.invoice_number && (
            <Button variant="ghost" size="sm" className="w-full text-red-600" onClick={() => setShowCancelInvoice(true)} disabled={busy}>
              송장 취소
            </Button>
          )}
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
      {/* 제품 교환 — 모든 진행/완료(취소 제외) 상태에서 가능 */}
      <Button variant="secondary" size="sm" className="w-full" onClick={() => setShowExchange(true)} disabled={busy}>
        <RefreshCw size={14} />
        제품 교환
      </Button>
      {order.exchanged_at && (
        <p className="text-[11px] text-emerald-600 px-1">✓ 교환 처리됨 — 아임웹 주문/결제는 그대로 유지됨</p>
      )}
    </div>
  );

  // ── 모달 (1벌) ──
  const modals = (
    <>
      {invoiceOpen && (
        <InvoiceModal open={invoiceOpen} onClose={() => setInvoiceOpen(false)} order={order} items={items} />
      )}
      {showExchange && (
        <OrderExchangeModal order={order} items={items} onClose={() => setShowExchange(false)} />
      )}
      <ConfirmModal
        open={showCancelOrder}
        onClose={() => setShowCancelOrder(false)}
        onConfirm={async () => { await cancelOrder.mutateAsync(order.id); }}
        title="주문 취소"
        message={cancelOrderMessage}
        confirmLabel="주문 취소"
        variant="danger"
      />
      <ConfirmModal
        open={showCancelInvoice}
        onClose={() => setShowCancelInvoice(false)}
        onConfirm={() => cancelInvoice.mutateAsync({ invNo: order.invoice_number!, orderId: order.id })}
        title="송장 취소"
        message={<>송장 <strong>{order.invoice_number}</strong>을 취소합니다.<br />ALPS 집하 전에만 가능합니다.</>}
        confirmLabel="송장 취소"
        variant="danger"
      />
      <ConfirmModal
        open={showPickup}
        onClose={() => setShowPickup(false)}
        onConfirm={async () => { await completePickup.mutateAsync(order.id); }}
        title="직접수령 완료"
        message={<>이 주문을 <strong>직접수령(대면 픽업)</strong>으로 완료합니다.<br />송장 없이 배송완료로 마감됩니다. 아임웹에서도 수령 처리해주세요.</>}
        confirmLabel="직접수령 완료"
      />
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

  return { hasPrimary, primary, secondary, modals };
}
