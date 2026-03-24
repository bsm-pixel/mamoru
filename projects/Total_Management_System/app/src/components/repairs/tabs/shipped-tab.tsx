'use client';

import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useUpdateRepairFields } from '@/hooks/use-repairs';
import { formatPhone, formatKRW, formatDateTime } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { Truck, AlertCircle, CreditCard, Check } from 'lucide-react';

interface ShippedTabProps {
  repairs: Repair[];
  isLoading: boolean;
  onSelect?: (id: string) => void;
}

/** 출고완료 탭: shipped + delivered + completed — 미납 시 [입금확인] */
export function ShippedTab({ repairs, isLoading, onSelect }: ShippedTabProps) {
  const updateFields = useUpdateRepairFields();

  const handleMarkPaid = (id: string) => {
    updateFields.mutate({ id, paid_at: new Date().toISOString() });
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
        <Truck size={32} className="mb-2 opacity-50" />
        <p className="text-sm">출고완료 건이 없습니다</p>
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
                {!r.paid_at && r.total_amount > 0 && (
                  <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-full text-[11px] font-medium bg-red-100 text-red-700">
                    <AlertCircle size={10} />
                    수리비 미납
                  </span>
                )}
                {r.paid_at && (
                  <span className="inline-flex items-center gap-0.5 px-1.5 py-0.5 rounded-full text-[11px] font-medium bg-green-100 text-green-700">
                    <Check size={10} />
                    입금완료
                  </span>
                )}
              </div>
              <div className="flex items-center gap-3 text-xs text-neutral-500">
                {r.invoice_number && (
                  <span className="font-mono">{r.invoice_number}</span>
                )}
                {r.shipped_at && (
                  <span>출고 {formatDateTime(r.shipped_at)}</span>
                )}
                {r.total_amount > 0 && (
                  <span className="font-medium text-neutral-600">{formatKRW(r.total_amount)}</span>
                )}
              </div>
            </div>
            {/* 미납 시 입금확인 버튼 */}
            {!r.paid_at && r.total_amount > 0 && (
              <Button
                variant="primary"
                size="sm"
                onClick={(e: React.MouseEvent) => { e.stopPropagation(); handleMarkPaid(r.id); }}
                loading={updateFields.isPending && updateFields.variables?.id === r.id}
                className="shrink-0"
              >
                <CreditCard size={12} />
                입금확인
              </Button>
            )}
          </div>
        </Card>
      ))}
    </div>
  );
}
