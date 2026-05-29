'use client';

import { useState } from 'react';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from './repair-status-badge';
import { RepairDetailCard } from './repair-detail-card';
import { InspectionForm } from './inspection-form';
import { InspectionSummary } from './inspection-summary';
import { SidebarActionCard } from './sidebar-action-card';
import { RepairTimeline } from './repair-timeline';
import { ReviewManagementCard } from '@/components/reviews/review-management-card';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { ClipboardCheck, ClipboardList } from 'lucide-react';
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
  const { data, isLoading, refetch } = useRepair(repairId);
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

      {/* 모바일: 내역작성 버튼 상단 배치 (스크롤 없이 즉시 접근) */}
      <div className="xl:hidden">
        <Button
          variant={inspections.length > 0 ? 'secondary' : 'primary'}
          size="sm"
          onClick={() => setShowInspection(true)}
          className="w-full"
        >
          {inspections.length > 0 ? (
            <><ClipboardCheck size={14} className="text-green-600" /> 내역작성완료 ({inspections.length}건)</>
          ) : (
            <><ClipboardList size={14} /> 내역작성하기</>
          )}
        </Button>
      </div>

      {/* 2컬럼 내부 레이아웃 */}
      <div className="flex flex-col xl:flex-row gap-4">
        {/* 메인 정보 */}
        <div className="flex-1 space-y-4 min-w-0">
          <RepairDetailCard repair={r} onUpdate={handleUpdate} />

          {/* PC: 검수 — 모달 트리거 버튼 */}
          <Button
            variant={inspections.length > 0 ? 'secondary' : 'primary'}
            size="sm"
            onClick={() => setShowInspection(true)}
            className="w-full hidden xl:flex"
          >
            {inspections.length > 0 ? (
              <><ClipboardCheck size={14} className="text-green-600" /> 내역작성완료 ({inspections.length}건)</>
            ) : (
              <><ClipboardList size={14} /> 내역작성하기</>
            )}
          </Button>

          {/* 검수 요약 (저장된 게 있을 때만) */}
          {inspections.length > 0 && <InspectionSummary inspections={inspections} />}

          {/* 검수 모달 */}
          <Modal
            open={showInspection}
            onClose={() => setShowInspection(false)}
            title="수리내역서 작성"
            className="max-w-2xl max-h-[90vh] overflow-y-auto"
            preventAutoClose
          >
            <InspectionForm
              repairId={r.id}
              existingInspections={inspections}
              totalScissors={r.qty_mamoru + r.qty_other}
              initialComment={r.admin_note || ''}
              onSaved={() => setShowInspection(false)}
            />
          </Modal>

        </div>

        {/* 사이드바 — 비용+액션+출고 통합 */}
        <div className="w-full xl:w-72 shrink-0 space-y-4">
          <SidebarActionCard repair={r} />
          {/* 067: 리뷰 관리 카드 — 취소 외 상태에서 표시 */}
          {r.status !== 'cancelled' && (
            <ReviewManagementCard
              source="repair"
              id={r.id}
              customerName={r.name}
              customerPhone={r.phone}
              promisedAt={(r as { review_promised_at?: string | null }).review_promised_at ?? null}
              promisedType={(r as { review_promised_type?: 'purchase' | 'repair' | 'consult' | null }).review_promised_type ?? null}
              promisedSubtype={(r as { review_promised_subtype?: string | null }).review_promised_subtype ?? null}
              requestSentAt={(r as { review_request_sent_at?: string | null }).review_request_sent_at ?? null}
              submittedAt={(r as { review_submitted_at?: string | null }).review_submitted_at ?? null}
              sourceType={(r as { proceed_type?: string | null }).proceed_type ?? null}
              onChanged={() => refetch()}
            />
          )}
          <RepairTimeline history={history} />
        </div>
      </div>
    </div>
  );
}
