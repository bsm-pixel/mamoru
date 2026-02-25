'use client';

import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { PipelineBar } from '@/components/dashboard/pipeline-bar';
import { UrgentList } from '@/components/dashboard/urgent-list';
import { Skeleton } from '@/components/ui/skeleton';
import { Button } from '@/components/ui/button';
import { useRepairDashboardStats } from '@/hooks/use-dashboard-stats';
import { useRepairs } from '@/hooks/use-repairs';
import { formatRelative } from '@/lib/utils/format';
import { AlertTriangle, ArrowRight } from 'lucide-react';

export default function RepairDashboardPage() {
  const router = useRouter();
  const { data: stats, isLoading: statsLoading } = useRepairDashboardStats();
  const { data: costData, isLoading: costLoading } = useRepairs({ status: 'cost_notified', limit: 5 });
  const { data: intakeData, isLoading: intakeLoading } = useRepairs({ status: 'intake', limit: 5 });

  return (
    <>
      <Topbar title="복원수리 대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-6">
        {/* 경과일 3일 이상 경고 배너 */}
        {!statsLoading && (stats?.staleCount || 0) > 0 && (
          <div className="flex items-center gap-2 bg-warning/10 border border-warning/30 rounded-lg px-4 py-3">
            <AlertTriangle size={18} className="text-warning shrink-0" />
            <p className="text-sm text-warning font-medium">
              검수/비용안내 후 <span className="font-bold">{stats?.staleCount}건</span>이 3일 이상 미처리 상태입니다
            </p>
          </div>
        )}

        {/* 파이프라인 */}
        <Card>
          <CardHeader>
            <CardTitle>복원수리 처리 현황</CardTitle>
          </CardHeader>
          {statsLoading ? (
            <Skeleton className="h-16" />
          ) : (
            <PipelineBar
              stages={(stats?.pipeline || []).map((s) => ({ ...s, href: `/repairs?status=${s.status}` }))}
            />
          )}
        </Card>

        {/* 긴급 리스트 2열 */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
          {/* 접수 처리 필요 */}
          {intakeLoading ? (
            <Skeleton className="h-40" />
          ) : (
            <UrgentList
              title="접수 처리 필요"
              items={(intakeData?.repairs || []).map((r) => ({
                id: r.id,
                label: r.name,
                sublabel: `${r.as_id || ''} · ${formatRelative(r.received_at)}`,
                badge: '접수',
                badgeColor: 'bg-info-soft text-info',
              }))}
              onItemClick={(id) => router.push(`/repairs?selected=${id}`)}
              emptyMessage="접수 대기 건 없음"
            />
          )}

          {/* 비용안내 대기 */}
          {costLoading ? (
            <Skeleton className="h-40" />
          ) : (
            <UrgentList
              title="비용안내 대기"
              items={(costData?.repairs || []).map((r) => ({
                id: r.id,
                label: r.name,
                sublabel: `${r.as_id || ''} · ${formatRelative(r.updated_at)}`,
                badge: '비용안내',
                badgeColor: 'bg-terracotta-soft/30 text-terracotta-deep',
              }))}
              onItemClick={(id) => router.push(`/repairs?selected=${id}`)}
              emptyMessage="비용안내 대기 건 없음"
            />
          )}
        </div>

        {/* 전체 목록 링크 */}
        <Button
          variant="ghost"
          className="w-full justify-center gap-1"
          onClick={() => router.push('/repairs')}
        >
          전체 복원수리 목록
          <ArrowRight size={14} />
        </Button>
      </div>
    </>
  );
}
