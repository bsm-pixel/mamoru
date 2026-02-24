'use client';

import { useState } from 'react';
import { useParams, useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from '@/components/repairs/repair-status-badge';
import { RepairDetailCard } from '@/components/repairs/repair-detail-card';
import { InspectionForm } from '@/components/repairs/inspection-form';
import { InspectionSummary } from '@/components/repairs/inspection-summary';
import { SidebarActionCard } from '@/components/repairs/sidebar-action-card';
import { RepairTimeline } from '@/components/repairs/repair-timeline';
import {
  useRepair,
  useUpdateRepairFields,
} from '@/hooks/use-repairs';
import { formatDateTime } from '@/lib/utils/format';
import type { RepairStatus } from '@/lib/supabase/types';
import { ArrowLeft } from 'lucide-react';

export default function RepairDetailPage() {
  const params = useParams();
  const router = useRouter();
  const id = params.id as string;
  const { data, isLoading } = useRepair(id);
  const updateFields = useUpdateRepairFields();
  const [showInspection, setShowInspection] = useState(false);

  if (isLoading) {
    return (
      <>
        <Topbar title="복원수리 상세" />
        <div className="px-4 md:px-6 py-4 space-y-4">
          <Skeleton className="h-8 w-32" />
          <Skeleton className="h-40 w-full" />
          <Skeleton className="h-40 w-full" />
        </div>
      </>
    );
  }

  if (!data) {
    return (
      <>
        <Topbar title="복원수리 상세" />
        <div className="flex items-center justify-center h-60 text-neutral-400">
          복원수리 건을 찾을 수 없습니다
        </div>
      </>
    );
  }

  const { repair: r, inspections, history } = data;
  const currentStatus = r.status as RepairStatus;
  const proceedType = r.proceed_type;

  const handleUpdate = async (fields: Record<string, unknown>) => {
    await updateFields.mutateAsync({ id: r.id, ...fields });
  };

  return (
    <>
      <Topbar title="복원수리 상세" />

      <div className="px-4 md:px-6 py-4 space-y-4 max-w-4xl">
        {/* 뒤로가기 */}
        <button
          onClick={() => router.back()}
          className="flex items-center gap-1 text-sm text-neutral-500 hover:text-indigo-black transition"
        >
          <ArrowLeft size={16} />
          목록으로
        </button>

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

        {/* 2컬럼 레이아웃 (PC) */}
        <div className="flex flex-col lg:flex-row gap-4">
          {/* 좌측: 메인 정보 */}
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

          {/* 우측: 사이드바 — 비용+액션+출고 통합 */}
          <div className="w-full lg:w-80 shrink-0 space-y-4">
            <SidebarActionCard repair={r} />
            <RepairTimeline history={history} />
          </div>
        </div>
      </div>
    </>
  );
}
