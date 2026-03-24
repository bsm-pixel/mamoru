'use client';

import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import {
  useUpdateRepairStatus,
  useSendRepairNotification,
} from '@/hooks/use-repairs';
import { formatPhone, formatRelative, formatKRW } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { CheckCircle, Send } from 'lucide-react';

interface InboundWaitingTabProps {
  repairs: Repair[];
  isLoading: boolean;
  onSelect?: (id: string) => void;
}

/** 입고대기 탭: 직접발송(confirmed_at + intake) + 방문수거(pickup_scheduled) — [입고 & 비용안내] */
export function InboundWaitingTab({ repairs, isLoading, onSelect }: InboundWaitingTabProps) {
  const updateStatus = useUpdateRepairStatus();
  const sendNotify = useSendRepairNotification();

  const handleCostNotice = (r: Repair) => {
    updateStatus.mutate(
      {
        id: r.id,
        status: 'cost_notified',
        service_cost: r.service_cost,
        shipping_fee: r.shipping_fee,
        total_amount: r.total_amount,
        note: `비용 안내: ${formatKRW(r.total_amount)}`,
      },
      {
        onSuccess: () => {
          sendNotify.mutate({
            repairId: r.id,
            template: 'as_cost_notice',
            extraData: {
              as_amount: String(r.service_cost),
              shipping_amount: String(r.shipping_fee),
              total_amount: String(r.total_amount),
            },
          });
        },
      }
    );
  };

  if (isLoading) {
    return (
      <div className="space-y-3 p-4">
        {[...Array(3)].map((_, i) => <Skeleton key={i} className="h-24 rounded-xl" />)}
      </div>
    );
  }

  if (!repairs.length) {
    return (
      <div className="flex flex-col items-center justify-center py-16 text-neutral-400">
        <CheckCircle size={32} className="mb-2 opacity-50" />
        <p className="text-sm">입고대기 건이 없습니다</p>
      </div>
    );
  }

  return (
    <div className="space-y-2 p-4">
      {repairs.map((r) => (
        <Card key={r.id} className="hover:bg-neutral-50 transition cursor-pointer" onClick={() => onSelect?.(r.id)}>
          <div className="flex items-start justify-between gap-3">
            <div className="flex-1 min-w-0">
              <div className="flex items-center gap-2 mb-1">
                <span className="text-sm font-semibold">{r.name}</span>
                <span className="text-xs text-neutral-500">{formatPhone(r.phone)}</span>
                <span className="px-1.5 py-0.5 rounded-full text-[11px] font-medium bg-info-soft text-info">
                  {r.proceed_type || '직접발송'}
                </span>
              </div>
              <div className="flex items-center gap-3 mt-1 text-xs text-neutral-600">
                {r.qty_mamoru > 0 && <span>마모루 {r.qty_mamoru}자루</span>}
                {r.qty_other > 0 && <span>타사 {r.qty_other}자루</span>}
                <span className="text-neutral-400">{formatRelative(r.received_at)}</span>
              </div>
            </div>
            <Button
              variant="primary"
              size="sm"
              onClick={(e: React.MouseEvent) => { e.stopPropagation(); handleCostNotice(r); }}
              loading={updateStatus.isPending && updateStatus.variables?.id === r.id}
              className="shrink-0"
            >
              <Send size={12} />
              입고 & 비용안내
            </Button>
          </div>
        </Card>
      ))}
    </div>
  );
}
