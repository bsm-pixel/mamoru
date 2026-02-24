'use client';

import { useState } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useSendRepairNotification, useUpdateRepairStatus } from '@/hooks/use-repairs';
import { formatKRW } from '@/lib/utils/format';
import { calcTotalCost } from '@/lib/repair/cost-calculator';
import type { Repair } from '@/lib/supabase/types';
import { Send, Calculator } from 'lucide-react';

interface CostSummaryProps {
  repair: Repair;
}

export function CostSummary({ repair }: CostSummaryProps) {
  const sendNotify = useSendRepairNotification();
  const updateStatus = useUpdateRepairStatus();
  const [editMode, setEditMode] = useState(false);
  const [serviceCost, setServiceCost] = useState(repair.service_cost);
  const [shippingFee, setShippingFee] = useState(repair.shipping_fee);

  const totalAmount = serviceCost + shippingFee;

  const handleRecalculate = () => {
    const costs = calcTotalCost(repair.qty_mamoru, repair.qty_other, repair.proceed_type);
    setServiceCost(costs.serviceCost);
    setShippingFee(costs.shippingFee);
  };

  const handleSendCostNotice = async () => {
    // 1) 비용 업데이트 + 상태 변경
    await updateStatus.mutateAsync({
      id: repair.id,
      status: 'cost_notified',
      service_cost: serviceCost,
      shipping_fee: shippingFee,
      total_amount: totalAmount,
      note: `비용 안내: ${formatKRW(totalAmount)}`,
    });

    // 2) 알림톡 발송
    sendNotify.mutate({
      repairId: repair.id,
      template: 'as_cost_notice',
      extraData: {
        as_amount: String(serviceCost),
        shipping_amount: String(shippingFee),
        total_amount: String(totalAmount),
      },
    });

    setEditMode(false);
  };

  const canSendNotice = ['intake', 'pickup_scheduled', 'picked_up', 'inspecting', 'cost_notified'].includes(repair.status);

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center justify-between">
          <span>비용 요약</span>
          {canSendNotice && (
            <Button variant="ghost" size="sm" onClick={() => setEditMode(!editMode)}>
              {editMode ? '취소' : '수정'}
            </Button>
          )}
        </CardTitle>
      </CardHeader>

      <dl className="grid grid-cols-[6rem_1fr] gap-y-2 text-sm">
        <dt className="text-neutral-500">수리비</dt>
        <dd>
          {editMode ? (
            <input
              type="number"
              value={serviceCost}
              onChange={(e) => setServiceCost(parseInt(e.target.value) || 0)}
              className="w-28 h-8 px-2 rounded border border-neutral-200 text-sm text-right"
            />
          ) : (
            <span className="font-medium">{formatKRW(repair.service_cost)}</span>
          )}
        </dd>
        <dt className="text-neutral-500">수거비</dt>
        <dd>
          {editMode ? (
            <input
              type="number"
              value={shippingFee}
              onChange={(e) => setShippingFee(parseInt(e.target.value) || 0)}
              className="w-28 h-8 px-2 rounded border border-neutral-200 text-sm text-right"
            />
          ) : (
            formatKRW(repair.shipping_fee)
          )}
        </dd>
        <dt className="text-neutral-500 font-semibold">합계</dt>
        <dd className="font-bold text-terracotta-deep">
          {editMode ? formatKRW(totalAmount) : formatKRW(repair.total_amount)}
        </dd>
      </dl>

      {editMode && (
        <div className="flex gap-2 mt-4">
          <Button variant="ghost" size="sm" onClick={handleRecalculate}>
            <Calculator size={14} />
            자동 계산
          </Button>
        </div>
      )}

      {/* 비용 안내 발송 */}
      {canSendNotice && (
        <div className="mt-4 pt-3 border-t border-neutral-100">
          <Button
            variant="primary"
            size="sm"
            onClick={handleSendCostNotice}
            loading={updateStatus.isPending || sendNotify.isPending}
            className="w-full"
          >
            <Send size={14} />
            {repair.status === 'cost_notified' ? '비용 안내 재발송' : '입고 & 비용안내'}
          </Button>
        </div>
      )}
    </Card>
  );
}
