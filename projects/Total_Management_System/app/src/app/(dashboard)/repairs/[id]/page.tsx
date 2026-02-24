'use client';

import { useState } from 'react';
import { useParams, useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { RepairStatusBadge } from '@/components/repairs/repair-status-badge';
import { RepairDetailCard } from '@/components/repairs/repair-detail-card';
import { InspectionForm } from '@/components/repairs/inspection-form';
import { InspectionSummary } from '@/components/repairs/inspection-summary';
import { CostSummary } from '@/components/repairs/cost-summary';
import { RepairTimeline } from '@/components/repairs/repair-timeline';
import {
  useRepair,
  useUpdateRepairStatus,
  useShipRepair,
  useCancelShipment,
} from '@/hooks/use-repairs';
import { getFilteredRepairTransitions, REPAIR_ACTION_LABEL } from '@/lib/repair/transitions';
import { formatDateTime } from '@/lib/utils/format';
import type { RepairStatus } from '@/lib/supabase/types';
import { ArrowLeft, Package, Truck, X } from 'lucide-react';

export default function RepairDetailPage() {
  const params = useParams();
  const router = useRouter();
  const id = params.id as string;
  const { data, isLoading } = useRepair(id);
  const updateStatus = useUpdateRepairStatus();
  const shipRepair = useShipRepair();
  const cancelShipment = useCancelShipment();
  const [showCancelConfirm, setShowCancelConfirm] = useState(false);
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
  const filtered = getFilteredRepairTransitions(currentStatus, proceedType);
  const busy = updateStatus.isPending;

  /** 상태별 액션 버튼 — 진행방식별 분기 */
  const renderActions = () => {
    if (currentStatus === 'completed' || currentStatus === 'cancelled') return null;

    return (
      <div className="flex flex-wrap gap-2">
        {/* cost_notified는 CostSummary에서 처리, shipped는 출고 섹션에서 처리 */}
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

        {/* 취소 */}
        {filtered.includes('cancelled' as RepairStatus) && (
          <Button
            variant="danger"
            size="sm"
            disabled={busy}
            onClick={() => setShowCancelConfirm(true)}
          >
            <X size={14} />
            취소
          </Button>
        )}
      </div>
    );
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
            <RepairDetailCard repair={r} />

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

            {/* 비용 요약 */}
            <CostSummary repair={r} />
          </div>

          {/* 우측: 사이드바 */}
          <div className="w-full lg:w-80 shrink-0 space-y-4">
            {/* 액션 */}
            <Card>
              <CardHeader>
                <CardTitle>액션</CardTitle>
              </CardHeader>
              {renderActions()}
            </Card>

            {/* 출고 섹션 — repairing 이상에서 표시 */}
            {(['repairing', 'shipped', 'delivered', 'completed'].includes(currentStatus)) && (
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

            {/* 이력 */}
            <RepairTimeline history={history} />
          </div>
        </div>
      </div>

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
