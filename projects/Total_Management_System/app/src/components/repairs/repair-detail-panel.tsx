'use client';

import { useState } from 'react';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from './repair-status-badge';
import { RepairDetailCard } from './repair-detail-card';
import { InspectionForm } from './inspection-form';
import { InspectionSummary } from './inspection-summary';
import { SidebarActionCard } from './sidebar-action-card';
import { RepairTimeline } from './repair-timeline';
import { RepairPrepSheetModal } from './repair-prep-sheet-modal';
import { ReviewManagementCard } from '@/components/reviews/review-management-card';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { ClipboardCheck, ClipboardList, Printer } from 'lucide-react';
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
  const [showPrep, setShowPrep] = useState(false);

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
    // @container: 내부 2단 배치를 '뷰포트' 아닌 '패널 폭' 기준으로 (좁은 상세패널에서 값 사라지던 깨짐 방지)
    <div className="@container flex-1 overflow-y-auto space-y-4 pr-1">
      {/* 헤더 — 이름·접수번호·상태 + 툴바(내역작성·준비표) 를 한 줄 정렬 (PC 툴바형) */}
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div className="min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <h2 className="text-lg font-bold truncate">{r.name}</h2>
            <RepairStatusBadge status={r.status} proceedType={proceedType} />
          </div>
          <p className="text-xs text-neutral-500 mt-0.5">
            {r.as_id} &middot; {formatDateTime(r.received_at)}
          </p>
        </div>
        {/* 툴바 — 컴팩트 버튼(풀폭 X) */}
        <div className="flex items-center gap-2 shrink-0">
          <Button
            variant={inspections.length > 0 ? 'secondary' : 'primary'}
            size="sm"
            onClick={() => setShowInspection(true)}
          >
            {inspections.length > 0 ? (
              <><ClipboardCheck size={14} className="text-green-600" /> 내역 {inspections.length}건</>
            ) : (
              <><ClipboardList size={14} /> 내역작성</>
            )}
          </Button>
          <button
            onClick={() => setShowPrep(true)}
            className="flex items-center gap-1 text-xs font-semibold text-neutral-600 border border-neutral-200 rounded-md px-2.5 py-1.5 hover:bg-neutral-50 transition"
          >
            <Printer size={12} /> 준비표
          </button>
        </div>
      </div>

      {showPrep && <RepairPrepSheetModal repairIds={[r.id]} onClose={() => setShowPrep(false)} />}

      {/* 2컬럼 내부 레이아웃 — 기본 정보 먼저(상단/좌), 액션은 바로 아래/우 (정보가 밀집돼 스크롤 부담 없음) */}
      <div className="flex flex-col @xl:flex-row gap-4">
        {/* 메인 정보 — 상단(좁을 때)/왼쪽(넓을 때) */}
        <div className="flex-1 space-y-4 min-w-0">
          <RepairDetailCard repair={r} onUpdate={handleUpdate} />

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
              onSaved={() => setShowInspection(false)}
            />
          </Modal>

        </div>

        {/* 사이드바(액션) — 정보 아래(좁을 때)/오른쪽(넓을 때) */}
        <div className="w-full @xl:w-72 shrink-0 space-y-4">
          <SidebarActionCard repair={r} />
          {/* 067: 리뷰 관리 카드 — 취소 외 상태에서 표시 (compact = 판매 패널과 동일 미니) */}
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
              compact={true}
              onChanged={() => refetch()}
            />
          )}
        </div>
      </div>

      {/* 이력 타임라인 — 참고용이라 항상 맨 아래 */}
      <RepairTimeline history={history} />
    </div>
  );
}
