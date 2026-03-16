'use client';

import { useState } from 'react';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from './repair-status-badge';
import { RepairDetailCard } from './repair-detail-card';
import { InspectionForm } from './inspection-form';
import { InspectionSummary } from './inspection-summary';
import { SidebarActionCard } from './sidebar-action-card';
import { RepairTimeline } from './repair-timeline';
import {
  useRepair,
  useUpdateRepairFields,
} from '@/hooks/use-repairs';
import { formatDateTime } from '@/lib/utils/format';
import type { RepairStatus } from '@/lib/supabase/types';
import { Scissors } from 'lucide-react';

interface RepairDetailPanelProps {
  repairId: string;
}

/** PC 우측 상세 패널 — repairs/page.tsx 마스터-디테일용 */
export function RepairDetailPanel({ repairId }: RepairDetailPanelProps) {
  const { data, isLoading } = useRepair(repairId);
  const updateFields = useUpdateRepairFields();
  const [showInspection, setShowInspection] = useState(false);

  if (isLoading) {
    return (
      <div className="space-y-4">
        <Skeleton className="h-8 w-48" />
        <Skeleton className="h-40 w-full" />
        <Skeleton className="h-40 w-full" />
      </div>
    );
  }

  if (!data) {
    return (
      <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
        <Scissors size={32} className="mb-2 opacity-50" />
        복원수리 건을 찾을 수 없습니다
      </div>
    );
  }

  const { repair: r, inspections, history } = data;
  const currentStatus = r.status as RepairStatus;
  const proceedType = r.proceed_type;

  const handleUpdate = async (fields: Record<string, unknown>) => {
    await updateFields.mutateAsync({ id: r.id, ...fields });
  };

  return (
    <div className="flex-1 overflow-y-auto space-y-4 pr-1">
      {/* 헤더 */}
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-lg font-bold">{r.name}</h2>
          <p className="text-xs text-neutral-500 mt-0.5">
            {r.as_id} &middot; {formatDateTime(r.received_at)}
          </p>
        </div>
        <RepairStatusBadge status={r.status} proceedType={proceedType} />
      </div>

      {/* 2컬럼 내부 레이아웃 */}
      <div className="flex flex-col xl:flex-row gap-4">
        {/* 메인 정보 */}
        <div className="flex-1 space-y-4 min-w-0">
          <RepairDetailCard repair={r} onUpdate={handleUpdate} />

          {/* 검수 섹션 */}
          {(showInspection || inspections.length > 0 ||
            ['inspecting', 'cost_notified', 'payment_confirmed', 'repairing', 'shipped', 'delivered', 'completed'].includes(currentStatus)
          ) && (
            <>
              <InspectionForm
                repairId={r.id}
                existingInspections={inspections}
                totalScissors={r.qty_mamoru + r.qty_other}
              />
              <InspectionSummary inspections={inspections} />
            </>
          )}
        </div>

        {/* 사이드바 — 비용+액션+출고 통합 */}
        <div className="w-full xl:w-72 shrink-0 space-y-4">
          <SidebarActionCard repair={r} />
          <RepairTimeline history={history} />
        </div>
      </div>
    </div>
  );
}
