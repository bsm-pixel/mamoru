'use client';

import { useState } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import {
  useUpdateRepairStatus,
  useUpdateRepairFields,
  useShipRepair,
  useCancelShipment,
  useSendRepairNotification,
} from '@/hooks/use-repairs';
import { getFilteredRepairTransitions, REPAIR_ACTION_LABEL } from '@/lib/repair/transitions';
import { formatKRW, formatDateTime } from '@/lib/utils/format';
import type { Repair, RepairStatus } from '@/lib/supabase/types';
import { Package, Truck, X, Send, CheckCircle, CreditCard } from 'lucide-react';
import { ConfirmModal } from '@/components/ui/confirm-modal';

interface SidebarActionCardProps {
  repair: Repair;
}

type ConfirmAction = 'cost_notice' | 'mark_paid' | 'mark_shipped' | 'cancel_shipment' | 'cancel_repair' | null;

export function SidebarActionCard({ repair: r }: SidebarActionCardProps) {
  const updateStatus = useUpdateRepairStatus();
  const updateFields = useUpdateRepairFields();
  const shipRepair = useShipRepair();
  const cancelShipment = useCancelShipment();
  const sendNotify = useSendRepairNotification();
  const [confirmAction, setConfirmAction] = useState<ConfirmAction>(null);
  const [paidNotify, setPaidNotify] = useState(true); // 입금확인 알림톡 발송 여부

  const currentStatus = r.status as RepairStatus;
  const proceedType = r.proceed_type;
  const filtered = getFilteredRepairTransitions(currentStatus, proceedType);
  const busy = updateStatus.isPending;
  const isTerminal = currentStatus === 'completed' || currentStatus === 'cancelled';

  // 비용안내 발송 가능 상태
  const canSendCostNotice = ['intake', 'pickup_scheduled', 'picked_up', 'inspecting', 'cost_notified'].includes(currentStatus);
  const isCostResend = currentStatus === 'cost_notified';

  // 입금확인 독립 버튼 표시 조건
  const canMarkPaid = !r.paid_at && ['cost_notified', 'repairing', 'ready_to_ship', 'shipped'].includes(currentStatus);

  // 출고완료 버튼 조건 (ready_to_ship 상태 + 송장 있음)
  const canMarkShipped = currentStatus === 'ready_to_ship' && !!r.invoice_number;

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

  // 입금확인 처리 (paid_at 설정 — 상태 변경 없음)
  const handleMarkPaid = () => {
    updateFields.mutate({
      id: r.id,
      paid_at: new Date().toISOString(),
      skip_notify: !paidNotify, // 체크 해제 시 알림톡 스킵
    });
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
              className="w-full"
            >
              <CreditCard size={14} />
              입금확인
            </Button>
          </div>
        ) : null}

        {/* 입고 & 비용안내 발송 */}
        {canSendCostNotice && (
          <div className="mt-3 pt-3 border-t border-neutral-100">
            <Button
              variant="primary"
              size="sm"
              onClick={() => setConfirmAction('cost_notice')}
              loading={updateStatus.isPending || sendNotify.isPending}
              className="w-full"
            >
              <Send size={14} />
              {isCostResend ? '비용 안내 재발송' : '입고 & 비용안내'}
            </Button>
          </div>
        )}

        {/* 액션 버튼 */}
        {!isTerminal && (
          <div className="mt-3 pt-3 border-t border-neutral-100 space-y-2">
            {filtered
              .filter((s) => s !== 'cancelled' && s !== 'cost_notified' && s !== 'shipped' && s !== 'repairing')
              .map((nextStatus) => (
                <Button
                  key={nextStatus}
                  variant="primary"
                  size="sm"
                  disabled={busy}
                  loading={updateStatus.variables?.status === nextStatus && busy}
                  onClick={() => updateStatus.mutate({ id: r.id, status: nextStatus })}
                  className="w-full"
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
                  className="w-full"
                >
                  <Truck size={14} />
                  출고완료
                </Button>
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
            </div>
          ) : (
            <Button
              variant="primary"
              size="sm"
              onClick={() => shipRepair.mutate({ id: r.id })}
              loading={shipRepair.isPending}
              className="w-full"
            >
              <Truck size={14} />
              송장 생성
            </Button>
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
      <ConfirmModal
        open={confirmAction === 'mark_paid'}
        onClose={() => { setConfirmAction(null); setPaidNotify(true); }}
        onConfirm={handleMarkPaid}
        title="입금 확인"
        message={
          <div className="space-y-3">
            <p>{r.name}님의 입금을 확인 처리합니다.<br />금액: {formatKRW(r.total_amount)}</p>
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
    </>
  );
}
