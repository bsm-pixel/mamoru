'use client';

import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useUpdateRepairStatus } from '@/hooks/use-repairs';
import { formatPhone, formatKRW } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { Package, Truck, AlertCircle } from 'lucide-react';

interface ReadyToShipTabProps {
  repairs: Repair[];
  isLoading: boolean;
  onSelect?: (id: string) => void;
}

/** 출고대기 탭: ready_to_ship — [출고완료] + 미입금 칩 */
export function ReadyToShipTab({ repairs, isLoading, onSelect }: ReadyToShipTabProps) {
  const updateStatus = useUpdateRepairStatus();

  const handleShipped = (id: string) => {
    updateStatus.mutate({
      id,
      status: 'shipped',
      shipped_at: new Date().toISOString(),
      note: '출고완료',
    });
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
        <Package size={32} className="mb-2 opacity-50" />
        <p className="text-sm">출고대기 건이 없습니다</p>
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
                <span className="text-xs text-neutral-400 font-mono">{r.as_id}</span>
                {!r.paid_at && (
                  <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-full text-[11px] font-medium bg-red-100 text-red-700">
                    <AlertCircle size={10} />
                    미입금
                  </span>
                )}
              </div>
              <div className="flex items-center gap-3 text-xs text-neutral-600">
                {r.invoice_number && (
                  <span className="font-mono text-neutral-500">{r.invoice_number}</span>
                )}
                {r.total_amount > 0 && (
                  <span className="font-medium">{formatKRW(r.total_amount)}</span>
                )}
              </div>
              {r.packed_at && (
                <span className="inline-flex items-center gap-0.5 mt-1 text-[11px] text-green-600 font-medium">
                  <Package size={10} />
                  포장완료
                </span>
              )}
            </div>
            <Button
              variant="primary"
              size="sm"
              onClick={(e: React.MouseEvent) => { e.stopPropagation(); handleShipped(r.id); }}
              loading={updateStatus.isPending && updateStatus.variables?.id === r.id}
              className="shrink-0"
            >
              <Truck size={12} />
              출고완료
            </Button>
          </div>
        </Card>
      ))}
    </div>
  );
}
