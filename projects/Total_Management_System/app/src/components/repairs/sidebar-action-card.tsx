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

interface SidebarActionCardProps {
  repair: Repair;
}

export function SidebarActionCard({ repair: r }: SidebarActionCardProps) {
  const updateStatus = useUpdateRepairStatus();
  const updateFields = useUpdateRepairFields();
  const shipRepair = useShipRepair();
  const cancelShipment = useCancelShipment();
  const sendNotify = useSendRepairNotification();
  const [showCancelConfirm, setShowCancelConfirm] = useState(false);

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
    await updateStatus.mutateAsync({
      id: r.id,
      status: 'cost_notified',
      service_cost: r.service_cost,
      shipping_fee: r.shipping_fee,
      total_amount: r.total_amount,
      note: `비용 안내: ${formatKRW(r.total_amount)}`,
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
              onClick={handleMarkPaid}
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
              onClick={handleSendCostNotice}
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
              .filter((s) => s !== 'cancelled' && s !== 'cost_notified' && s !== 'shipped')
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
              <Button
                variant="danger"
                size="sm"
                disabled={busy}
                onClick={() => setShowCancelConfirm(true)}
                className="w-full"
              >
                <X size={14} />
                취소
              </Button>
            )}
          </div>
        )}
      </Card>

      {/* 출고 섹션 — repairing 이상 */}
      {(['repairing', 'ready_to_ship', 'shipped', 'delivered', 'completed'].includes(currentStatus)) && (
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
                  onClick={handleMarkShipped}
                  loading={updateStatus.isPending}
                  className="w-full"
                >
                  <Truck size={14} />
                  출고완료
                </Button>
              )}
              {/* 송장 취소 (ready_to_ship에서만) */}
              {currentStatus === 'ready_to_ship' && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => cancelShipment.mutate({ id: r.id })}
                  loading={cancelShipment.isPending}
                  className="text-error"
                >
                  송장 취소
                </Button>
              )}
              {/* 배송중 상태에서 송장 취소 */}
              {currentStatus === 'shipped' && (
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => cancelShipment.mutate({ id: r.id })}
                  loading={cancelShipment.isPending}
                  className="text-error"
                >
                  송장 취소
                </Button>
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

      {/* 취소 확인 모달 */}
      {showCancelConfirm && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40">
          <div className="bg-white rounded-xl shadow-xl p-6 w-full max-w-sm mx-4">
            <h3 className="text-lg font-bold text-neutral-900">복원수리 취소</h3>
            <p className="text-sm text-neutral-600 mt-2">
              정말 이 복원수리 접수를 취소하시겠습니까?
            </p>
            <div className="flex gap-2 mt-5 justify-end">
              <Button variant="ghost" size="sm" onClick={() => setShowCancelConfirm(false)}>
                돌아가기
              </Button>
              <Button
                variant="danger"
                size="sm"
                loading={updateStatus.isPending}
                onClick={() => {
                  updateStatus.mutate(
                    { id: r.id, status: 'cancelled', note: '관리자 취소' },
                    { onSettled: () => setShowCancelConfirm(false) }
                  );
                }}
              >
                취소 확정
              </Button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}
