'use client';

import { useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import {
  useUpdateRepairStatus,
  useUpdateRepairFields,
  useShipRepair,
  useCancelShipment,
  useRecallRepairPickup,
  useSendRepairNotification,
  useDeleteRepair,
} from '@/hooks/use-repairs';
import { invalidateFinancialQueries } from '@/lib/query/invalidate-keys';
import { getFilteredRepairTransitions, REPAIR_ACTION_LABEL } from '@/lib/repair/transitions';
import { formatKRW, formatDateTime } from '@/lib/utils/format';
import type { Repair, RepairStatus } from '@/lib/supabase/types';
import { Package, Truck, X, Send, CheckCircle, CreditCard } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { MergedShipModal } from './merged-ship-modal';

interface SidebarActionCardProps {
  repair: Repair;
}

type ConfirmAction = 'cost_notice' | 'mark_paid' | 'mark_shipped' | 'cancel_shipment' | 'cancel_repair' | 'delete_repair' | 'visit_checkout' | null;
type PayMethod = 'transfer' | 'card' | 'cash';
const PAY_METHODS: { id: PayMethod; label: string }[] = [
  { id: 'transfer', label: '이체' },
  { id: 'card', label: '카드' },
  { id: 'cash', label: '현금' },
];

export function SidebarActionCard({ repair: r }: SidebarActionCardProps) {
  const queryClient = useQueryClient();
  const updateStatus = useUpdateRepairStatus();
  const updateFields = useUpdateRepairFields();
  const shipRepair = useShipRepair();
  const cancelShipment = useCancelShipment();
  const recallPickup = useRecallRepairPickup();
  const sendNotify = useSendRepairNotification();
  const deleteRepair = useDeleteRepair();
  const [confirmAction, setConfirmAction] = useState<ConfirmAction>(null);
  const [mergedShipOpen, setMergedShipOpen] = useState(false);
  const [paidNotify, setPaidNotify] = useState(true); // 입금확인 알림톡 발송 여부
  const [payMethod, setPayMethod] = useState<PayMethod>('transfer'); // 120: 결제수단
  const [visitPaid, setVisitPaid] = useState(true); // 직접방문 현장결제 시 입금완료 처리 여부
  const isDirectVisit = r.proceed_type === '직접방문';

  const currentStatus = r.status as RepairStatus;
  const proceedType = r.proceed_type;
  const filtered = getFilteredRepairTransitions(currentStatus, proceedType);
  const busy = updateStatus.isPending;
  const isTerminal = currentStatus === 'completed' || currentStatus === 'cancelled';

  // 비용안내 발송 가능 상태
  const canSendCostNotice = ['intake', 'pickup_scheduled', 'picked_up', 'inspecting', 'cost_notified'].includes(currentStatus);
  const isCostResend = currentStatus === 'cost_notified';

  // 입금확인 독립 버튼 표시 조건 — 미입금이면 배송완료/완료 이후에도 가능
  //   (ALPS 자동추적으로 shipped→delivered→completed 전이 후 입금확인 길이 막히던 버그 fix 2026-06-11)
  const canMarkPaid = !r.paid_at && ['cost_notified', 'repairing', 'ready_to_ship', 'shipped', 'delivered', 'completed'].includes(currentStatus);

  // 출고완료 버튼 조건 (ready_to_ship 상태 + 송장 있음)
  const canMarkShipped = currentStatus === 'ready_to_ship' && !!r.invoice_number;

  // 합포장 출고 버튼 조건 (송장 없을 때 — 다른 주문 송장에 합쳐 발송한 케이스)
  const canMergedShip = currentStatus === 'ready_to_ship' && !r.invoice_number;

  const handleSendCostNotice = async () => {
    const isFree = r.total_amount === 0; // 무상 처리 여부

    // 비용안내 → 자동으로 repairing 전환 (작업시작 별도 클릭 불필요)
    await updateStatus.mutateAsync({
      id: r.id,
      status: 'cost_notified',
      service_cost: r.service_cost,
      shipping_fee: r.shipping_fee,
      total_amount: r.total_amount,
      note: isFree ? '무상 처리 비용 안내' : `비용 안내: ${formatKRW(r.total_amount)}`,
    });

    // 0원이면 자동 입금완료 처리
    if (isFree && !r.paid_at) {
      await updateFields.mutateAsync({
        id: r.id,
        paid_at: new Date().toISOString(),
      });
    }

    // 비용안내 발송 후 자동으로 수리중 전환
    await updateStatus.mutateAsync({
      id: r.id,
      status: 'repairing',
      note: '비용안내 후 자동 작업시작',
    });
    sendNotify.mutate({
      repairId: r.id,
      template: 'as_cost_notice',
      extraData: {
        as_amount: String(r.service_cost),
        shipping_amount: String(r.shipping_fee),
        total_amount: String(r.total_amount),
      },
    });
  };

  // 입금확인 처리 (paid_at 설정 — 상태 변경 없음) + 결제수단 기록
  const handleMarkPaid = () => {
    updateFields.mutate({
      id: r.id,
      paid_at: new Date().toISOString(),
      payment_method: payMethod, // 120
      skip_notify: !paidNotify, // 체크 해제 시 알림톡 스킵
    });
  };

  // 직접방문 현장결제: 비용확정→작업시작(알림톡 없음) + 결제수단 + 입금(옵션) 한 번에
  const handleVisitCheckout = async () => {
    if (currentStatus === 'intake') {
      await updateStatus.mutateAsync({
        id: r.id,
        status: 'cost_notified',
        service_cost: r.service_cost,
        shipping_fee: r.shipping_fee,
        total_amount: r.total_amount,
        note: `방문 현장 비용확정: ${formatKRW(r.total_amount)}`,
      });
    }
    await updateStatus.mutateAsync({ id: r.id, status: 'repairing', note: '방문 현장 작업시작' });
    await updateFields.mutateAsync({
      id: r.id,
      payment_method: payMethod,
      ...(visitPaid && !r.paid_at ? { paid_at: new Date().toISOString(), skip_notify: true } : {}),
    });
    setConfirmAction(null);
  };

  // 출고완료 처리 (ready_to_ship → shipped + shipped_at 설정)
  const handleMarkShipped = () => {
    updateStatus.mutate({
      id: r.id,
      status: 'shipped',
      shipped_at: new Date().toISOString(),
      note: '출고완료',
    });
  };

  // 합포장 출고 성공 시 캐시 갱신 (모달 자체 API 호출 → 부모에서 invalidate)
  const handleMergedShipSuccess = () => {
    queryClient.invalidateQueries({ queryKey: ['repair', r.id] });
    queryClient.invalidateQueries({ queryKey: ['repair-tabs'] });
    invalidateFinancialQueries(queryClient);
  };

  return (
    <>
      <Card>
        {/* 비용 요약 (compact) */}
        <CardHeader>
          <CardTitle>비용</CardTitle>
        </CardHeader>
        <div className="text-sm space-y-1">
          <div className="flex justify-between">
            <span className="text-neutral-500">수리비</span>
            <span>{formatKRW(r.service_cost)}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">수거비</span>
            <span>{formatKRW(r.shipping_fee)}</span>
          </div>
          <div className="flex justify-between font-bold border-t border-neutral-100 pt-1 mt-1">
            <span>합계</span>
            <span className="text-terracotta-deep">{formatKRW(r.total_amount)}</span>
          </div>
        </div>

        {/* 입금 상태 표시 */}
        {r.paid_at ? (
          <div className="mt-3 pt-3 border-t border-neutral-100">
            <div className="flex items-center gap-1.5 text-sm text-success font-medium">
              <CheckCircle size={14} />
              입금완료
              <span className="text-xs text-neutral-400 font-normal ml-auto">
                {formatDateTime(r.paid_at)}
              </span>
            </div>
          </div>
        ) : canMarkPaid ? (
          <div className="mt-3 pt-3 border-t border-neutral-100">
            <Button
              variant="primary"
              size="sm"
              onClick={() => setConfirmAction('mark_paid')}
              loading={updateFields.isPending}
            >
              <CreditCard size={14} />
              입금확인
            </Button>
          </div>
        ) : null}

        {/* 입고 & 비용안내 발송 (택배·방문수거) / 방문 현장결제 (직접방문 — 알림톡 없음) */}
        {canSendCostNotice && (
          <div className="mt-3 pt-3 border-t border-neutral-100">
            {isDirectVisit ? (
              <Button
                variant="primary"
                size="sm"
                onClick={() => setConfirmAction('visit_checkout')}
                loading={updateStatus.isPending || updateFields.isPending}
              >
                🏪 방문 확정 · 현장결제
              </Button>
            ) : (
              <Button
                variant="primary"
                size="sm"
                onClick={() => setConfirmAction('cost_notice')}
                loading={updateStatus.isPending || sendNotify.isPending}
              >
                <Send size={14} />
                {isCostResend ? '비용 안내 재발송' : '입고 & 비용안내'}
              </Button>
            )}
          </div>
        )}

        {/* 액션 버튼 */}
        {!isTerminal && (
          <div className="mt-3 pt-3 border-t border-neutral-100 space-y-2">
            {filtered
              .filter((s) => s !== 'cancelled' && s !== 'cost_notified' && s !== 'shipped' && s !== 'repairing' && s !== 'delivered')
              .map((nextStatus) => (
                <Button
                  key={nextStatus}
                  variant="primary"
                  size="sm"
                  disabled={busy}
                  loading={updateStatus.variables?.status === nextStatus && busy}
                  onClick={() => updateStatus.mutate({ id: r.id, status: nextStatus })}
                >
                  {REPAIR_ACTION_LABEL[nextStatus]}
                </Button>
              ))}

            {filtered.includes('cancelled' as RepairStatus) && (
              <button
                disabled={busy}
                onClick={() => setConfirmAction('cancel_repair')}
                className="w-full text-center text-xs text-red-400 hover:text-red-600 py-1.5 transition disabled:opacity-50"
              >
                취소
              </button>
            )}
            <button
              disabled={busy}
              onClick={() => setConfirmAction('delete_repair')}
              className="w-full text-center text-xs text-neutral-400 hover:text-red-500 py-1 transition disabled:opacity-50"
            >
              삭제
            </button>
          </div>
        )}
      </Card>

      {/* 출고 섹션 — cost_notified 이상 (입금확인 후 바로 송장생성 가능) */}
      {(['cost_notified', 'repairing', 'ready_to_ship', 'shipped', 'delivered', 'completed'].includes(currentStatus)) && (
        <Card>
          <CardHeader>
            <CardTitle>
              <Truck size={16} className="inline mr-1.5" />
              출고
            </CardTitle>
          </CardHeader>
          {r.invoice_number ? (
            <div className="space-y-2">
              <div className="flex items-center gap-2 text-sm">
                <Package size={14} className="text-success" />
                <span className="font-mono font-medium">{r.invoice_number}</span>
              </div>
              <p className="text-xs text-neutral-500">{r.courier_name || '롯데택배'}</p>
              {r.shipped_at && (
                <p className="text-xs text-neutral-400">발송: {formatDateTime(r.shipped_at)}</p>
              )}
              {/* 출고완료 버튼 (ready_to_ship → shipped) */}
              {canMarkShipped && (
                <Button
                  variant="primary"
                  size="sm"
                  onClick={() => setConfirmAction('mark_shipped')}
                  loading={updateStatus.isPending}
                >
                  <Truck size={14} />
                  출고완료
                </Button>
              )}
              {/* shipped 상태 — ALPS 자동 추적 안내 + 수동 fallback */}
              {currentStatus === 'shipped' && (
                <div className="pt-1 space-y-1">
                  <p className="text-xs text-neutral-400 leading-relaxed">
                    ALPS 인수자등록 자동 감지 시 배송완료 전환됩니다 (1시간마다 자동 확인)
                  </p>
                  <button
                    onClick={() =>
                      updateStatus.mutate({
                        id: r.id,
                        status: 'delivered',
                        delivered_at: new Date().toISOString(),
                        note: '수동 배송완료 처리 (ALPS 추적 fallback)',
                      })
                    }
                    disabled={updateStatus.isPending}
                    className="text-xs text-neutral-500 hover:text-neutral-700 underline transition disabled:opacity-50"
                  >
                    수동 배송완료 처리
                  </button>
                </div>
              )}
              {/* 송장 취소 */}
              {['ready_to_ship', 'shipped'].includes(currentStatus) && (
                <button
                  onClick={() => setConfirmAction('cancel_shipment')}
                  disabled={cancelShipment.isPending}
                  className="text-xs text-red-400 hover:text-red-600 py-1 transition disabled:opacity-50"
                >
                  {cancelShipment.isPending ? '취소 중...' : '송장 취소'}
                </button>
              )}
              {/* 정밀 재점검 재수거 (출고된 건 회수 — 롯데 반품 API 02). 알림톡 없음, 취소는 ALPS 수동 */}
              {['shipped', 'delivered', 'completed'].includes(currentStatus) && (
                r.recall_invoice_number ? (
                  <div className="mt-1 text-xs text-blue-700 bg-blue-50 rounded-lg px-2.5 py-2">
                    ↩ 재수거 송장 <b className="font-mono">{r.recall_invoice_number}</b>
                    {r.recall_booked_at && <span className="text-neutral-400 ml-1">({formatDateTime(r.recall_booked_at)})</span>}
                    <p className="text-[11px] text-neutral-400 mt-0.5">접수 후 취소는 ALPS에서 수동입니다.</p>
                  </div>
                ) : (
                  <button
                    onClick={() => { if (window.confirm('정밀 재점검 재수거를 접수합니다 (고객집 방문수거).\n접수 후 취소는 ALPS에서 수동입니다. 진행할까요?')) recallPickup.mutate(r.id); }}
                    disabled={recallPickup.isPending}
                    className="w-full flex items-center justify-center gap-1.5 py-2 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 transition disabled:opacity-50"
                  >
                    <Truck size={13} /> {recallPickup.isPending ? '접수 중…' : '정밀 재점검 재수거 접수'}
                  </button>
                )
              )}
            </div>
          ) : (
            <div className="space-y-2">
              <Button
                variant="primary"
                size="sm"
                onClick={() => shipRepair.mutate({ id: r.id })}
                loading={shipRepair.isPending}
              >
                <Truck size={14} />
                송장 생성
              </Button>
              {canMergedShip && (
                <Button
                  variant="secondary"
                  size="sm"
                  onClick={() => setMergedShipOpen(true)}
                >
                  <Package size={14} />
                  판매건 합포장 출고
                </Button>
              )}
              {/* 119: 택배 없이 매장에서 직접 전달(직접수령) → 바로 완료 처리. 작업중/출고대기에서만 */}
              {['repairing', 'ready_to_ship'].includes(currentStatus) && (
                <button
                  onClick={() => {
                    if (!window.confirm(`${r.name}님께 매장에서 직접 전달(택배 없이) 완료로 처리합니다.\n계속할까요?`)) return;
                    updateStatus.mutate({
                      id: r.id,
                      status: 'delivered',
                      delivered_at: new Date().toISOString(),
                      delivery_method: 'pickup',
                      note: '직접 수령 (매장 전달)',
                    });
                  }}
                  disabled={updateStatus.isPending}
                  className="w-full text-xs font-semibold text-neutral-600 border border-neutral-200 rounded-lg py-2 hover:bg-neutral-50 transition disabled:opacity-50"
                >
                  🏪 직접 수령 (매장 전달) — 택배 없이 완료
                </button>
              )}
            </div>
          )}
        </Card>
      )}

      {/* 확인 모달들 */}
      <ConfirmModal
        open={confirmAction === 'cost_notice'}
        onClose={() => setConfirmAction(null)}
        onConfirm={handleSendCostNotice}
        title="비용 안내 발송"
        message={<>고객에게 <strong>비용 안내 알림톡</strong>이 발송됩니다.<br />금액: {formatKRW(r.total_amount)}{r.total_amount === 0 ? ' (무상 처리)' : ''}</>}
        confirmLabel="발송"
      />
      {/* 직접방문 현장결제 — 비용확정 + 결제수단 + 입금완료 한 번에 (알림톡 없음) */}
      <ConfirmModal
        open={confirmAction === 'visit_checkout'}
        onClose={() => { setConfirmAction(null); setVisitPaid(true); }}
        onConfirm={handleVisitCheckout}
        title="방문 확정 · 현장결제"
        message={
          <div className="space-y-3">
            <div className="text-sm text-neutral-600 bg-neutral-50 rounded-lg p-2.5 leading-relaxed">
              <div>수량: 마모루 {r.qty_mamoru}자루{r.qty_other > 0 ? ` · 타사 ${r.qty_other}자루` : ''}</div>
              <div className="font-semibold text-neutral-900 mt-0.5">금액: {formatKRW(r.total_amount)}{r.total_amount === 0 ? ' (무상)' : ''}</div>
              <div className="text-xs text-neutral-400 mt-1">※ 수량이 실제와 다르면 접수정보에서 먼저 조정하세요</div>
            </div>
            <div>
              <p className="text-xs text-neutral-500 mb-1.5">결제수단</p>
              <div className="flex gap-2">
                {PAY_METHODS.map((m) => (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setPayMethod(m.id)}
                    className={`flex-1 py-1.5 rounded-lg text-sm font-semibold border transition ${payMethod === m.id ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-600 border-neutral-200 hover:bg-neutral-50'}`}
                  >
                    {m.label}
                  </button>
                ))}
              </div>
            </div>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={visitPaid}
                onChange={(e) => setVisitPaid(e.target.checked)}
                className="w-4 h-4 rounded border-neutral-300 text-terracotta focus:ring-terracotta"
              />
              <span className="text-sm text-neutral-600">입금 완료로 처리 (현장 결제)</span>
            </label>
            <p className="text-xs text-neutral-400">알림톡은 발송하지 않습니다 (고객 현장 방문).</p>
          </div>
        }
        confirmLabel="현장 처리 완료"
      />
      <ConfirmModal
        open={confirmAction === 'mark_paid'}
        onClose={() => { setConfirmAction(null); setPaidNotify(true); }}
        onConfirm={handleMarkPaid}
        title="입금 확인"
        message={
          <div className="space-y-3">
            <p>{r.name}님의 입금을 확인 처리합니다.<br />금액: {formatKRW(r.total_amount)}</p>
            <div>
              <p className="text-xs text-neutral-500 mb-1.5">결제수단</p>
              <div className="flex gap-2">
                {PAY_METHODS.map((m) => (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => setPayMethod(m.id)}
                    className={`flex-1 py-1.5 rounded-lg text-sm font-semibold border transition ${payMethod === m.id ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-600 border-neutral-200 hover:bg-neutral-50'}`}
                  >
                    {m.label}
                  </button>
                ))}
              </div>
            </div>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="checkbox"
                checked={paidNotify}
                onChange={(e) => setPaidNotify(e.target.checked)}
                className="w-4 h-4 rounded border-neutral-300 text-terracotta focus:ring-terracotta"
              />
              <span className="text-sm text-neutral-600">고객에게 입금확인 알림톡 발송</span>
            </label>
            {!paidNotify && (
              <p className="text-xs text-neutral-400">알림톡 없이 내부 입금완료 처리만 합니다</p>
            )}
          </div>
        }
        confirmLabel={paidNotify ? '입금 확인 + 알림톡' : '입금완료 표시'}
      />
      <ConfirmModal
        open={confirmAction === 'mark_shipped'}
        onClose={() => setConfirmAction(null)}
        onConfirm={handleMarkShipped}
        title="출고 완료"
        message={<>송장 {r.invoice_number}으로 출고 완료 처리합니다.<br />고객에게 <strong>출고 알림톡</strong>이 자동 발송됩니다.</>}
        confirmLabel="출고 완료"
      />
      <MergedShipModal
        open={mergedShipOpen}
        onClose={() => setMergedShipOpen(false)}
        repairId={r.id}
        onSuccess={handleMergedShipSuccess}
      />
      <ConfirmModal
        open={confirmAction === 'cancel_shipment'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => cancelShipment.mutateAsync({ id: r.id })}
        title="송장 취소"
        message="생성된 송장을 취소합니다. 롯데택배 집하 전에만 가능합니다."
        confirmLabel="송장 취소"
        variant="danger"
      />
      <ConfirmModal
        open={confirmAction === 'cancel_repair'}
        onClose={() => setConfirmAction(null)}
        onConfirm={() => updateStatus.mutateAsync({ id: r.id, status: 'cancelled', note: '관리자 취소' })}
        title="복원수리 취소"
        message="정말 이 복원수리 접수를 취소하시겠습니까?"
        confirmLabel="취소 확정"
        variant="danger"
      />
      <ConfirmModal
        open={confirmAction === 'delete_repair'}
        onClose={() => setConfirmAction(null)}
        onConfirm={async () => { await deleteRepair.mutateAsync(r.id); }}
        title="복원수리 삭제"
        message={<>이 건을 <strong>완전히 삭제</strong>합니다.<br />알림톡은 발송되지 않습니다. 복구할 수 없습니다.</>}
        confirmLabel="삭제"
        variant="danger"
      />
    </>
  );
}
