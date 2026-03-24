'use client';

import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useUpdateRepairStatus } from '@/hooks/use-repairs';
import { formatPhone, formatRelative } from '@/lib/utils/format';
import type { Repair } from '@/lib/supabase/types';
import { CheckCircle, MapPin } from 'lucide-react';

interface PickupNeededTabProps {
  repairs: Repair[];
  isLoading: boolean;
  onSelect?: (id: string) => void;
}

/** 수거접수필요 탭: 방문수거 + confirmed_at IS NULL 또는 접수확인 완료 — [수거접수 완료] */
export function PickupNeededTab({ repairs, isLoading, onSelect }: PickupNeededTabProps) {
  const updateStatus = useUpdateRepairStatus();

  const handlePickupDone = (id: string) => {
    updateStatus.mutate({ id, status: 'pickup_scheduled', note: '수거접수 완료' });
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
        <p className="text-sm">수거접수 대기 건이 없습니다</p>
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
              </div>
              <div className="flex items-center gap-1 text-xs text-neutral-500">
                <MapPin size={11} />
                <span className="truncate">{r.address} {r.address_detail}</span>
              </div>
              <div className="flex items-center gap-3 mt-1.5 text-xs text-neutral-600">
                {r.qty_mamoru > 0 && <span>마모루 {r.qty_mamoru}자루</span>}
                {r.qty_other > 0 && <span>타사 {r.qty_other}자루</span>}
                <span className="text-neutral-400">{formatRelative(r.received_at)}</span>
              </div>
            </div>
            <Button
              variant="primary"
              size="sm"
              onClick={(e: React.MouseEvent) => { e.stopPropagation(); handlePickupDone(r.id); }}
              loading={updateStatus.isPending && updateStatus.variables?.id === r.id}
              className="shrink-0"
            >
              수거접수 완료
            </Button>
          </div>
        </Card>
      ))}
    </div>
  );
}
