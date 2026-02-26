'use client';

import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { StatsCard } from '@/components/dashboard/stats-card';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { Button } from '@/components/ui/button';
import { useConsultationDashboardStats } from '@/hooks/use-dashboard-stats';
import {
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
  CONSULTATION_TYPE_LABEL,
} from '@/lib/utils/format';
import { Inbox, Loader, CheckCircle, ArrowRight } from 'lucide-react';

export default function ConsultationDashboardPage() {
  const router = useRouter();
  const { data: stats, isLoading } = useConsultationDashboardStats();

  return (
    <>
      <Topbar title="상담 대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-6">
        {/* R2: 통계 카드 3개 */}
        {isLoading ? (
          <div className="grid grid-cols-3 gap-3">
            {Array.from({ length: 3 }).map((_, i) => (
              <Skeleton key={i} className="h-20" />
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-3 gap-3">
            <StatsCard
              label="신규접수"
              value={stats?.newIntake || 0}
              icon={Inbox}
              color="text-info"
              bgColor="bg-info/10"
              subtitle="6시간 이내"
            />
            <StatsCard
              label="진행중"
              value={stats?.inProgress || 0}
              icon={Loader}
              color="text-terracotta"
              bgColor="bg-terracotta/10"
              subtitle="6시간 이후"
            />
            <StatsCard
              label="상담완료"
              value={stats?.completedMonth || 0}
              icon={CheckCircle}
              color="text-success"
              bgColor="bg-success/10"
              subtitle="최근 1달"
            />
          </div>
        )}

        {/* 오늘 일정 타임라인 */}
        <Card>
          <CardHeader>
            <CardTitle>오늘 일정</CardTitle>
          </CardHeader>
          {isLoading ? (
            <Skeleton className="h-32" />
          ) : !stats?.todaySchedule?.length ? (
            <div className="text-center py-6">
              <p className="text-sm text-neutral-400">오늘 예정된 일정이 없습니다</p>
            </div>
          ) : (
            <div className="divide-y divide-neutral-100 -mx-5">
              {stats.todaySchedule.map((c) => (
                <button
                  key={c.id}
                  type="button"
                  onClick={() => router.push(`/consultations/${c.id}`)}
                  className="w-full flex items-center justify-between px-5 py-3 text-left hover:bg-warm-ivory/60 transition"
                >
                  <div className="flex items-center gap-3 min-w-0">
                    <span className="text-sm font-mono font-bold text-terracotta w-12 shrink-0">
                      {c.visit_time || '--:--'}
                    </span>
                    <div className="min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="text-sm font-semibold truncate">{c.name}</span>
                        <Badge className={CONSULTATION_STATUS_COLOR[c.status] || ''}>
                          {CONSULTATION_STATUS_LABEL[c.status]}
                        </Badge>
                      </div>
                      <p className="text-xs text-neutral-500 mt-0.5">
                        {CONSULTATION_TYPE_LABEL[c.consultation_type] || c.consultation_type}
                      </p>
                    </div>
                  </div>
                </button>
              ))}
            </div>
          )}
        </Card>

        {/* 전체 목록 링크 */}
        <Button
          variant="ghost"
          className="w-full justify-center gap-1"
          onClick={() => router.push('/consultations')}
        >
          전체 상담 목록
          <ArrowRight size={14} />
        </Button>
      </div>
    </>
  );
}
