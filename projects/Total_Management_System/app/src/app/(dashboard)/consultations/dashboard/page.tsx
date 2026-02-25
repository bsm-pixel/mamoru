'use client';

import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { StatsCard } from '@/components/dashboard/stats-card';
import { UrgentList } from '@/components/dashboard/urgent-list';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { Button } from '@/components/ui/button';
import { useConsultationDashboardStats } from '@/hooks/use-dashboard-stats';
import { useConsultations } from '@/hooks/use-consultations';
import {
  formatRelative,
  CONSULTATION_STATUS_LABEL,
  CONSULTATION_STATUS_COLOR,
  CONSULTATION_TYPE_LABEL,
} from '@/lib/utils/format';
import { AlertCircle, Store, MapPin, Clock, CalendarDays, ArrowRight } from 'lucide-react';

export default function ConsultationDashboardPage() {
  const router = useRouter();
  const { data: stats, isLoading: statsLoading } = useConsultationDashboardStats();
  const { data: urgentData, isLoading: urgentLoading } = useConsultations({
    statuses: ['pending_admin'],
    limit: 5,
  });

  return (
    <>
      <Topbar title="상담 대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-6">
        {/* 통계 카드 */}
        {statsLoading ? (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {Array.from({ length: 4 }).map((_, i) => (
              <Skeleton key={i} className="h-20" />
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <StatsCard label="미확인" value={stats?.pendingAdmin || 0} icon={AlertCircle} color="text-warning" bgColor="bg-warning/10" />
            <StatsCard label="오늘 매장" value={stats?.todayStore || 0} icon={Store} color="text-info" bgColor="bg-info/10" />
            <StatsCard label="오늘 출장" value={stats?.todayField || 0} icon={MapPin} color="text-terracotta" bgColor="bg-terracotta/10" />
            <StatsCard label="제안 대기" value={stats?.suggested || 0} icon={Clock} color="text-neutral-600" bgColor="bg-neutral-100" />
          </div>
        )}

        {/* 이번주 예약 */}
        {!statsLoading && (
          <div className="flex items-center gap-2 px-1">
            <CalendarDays size={16} className="text-terracotta" />
            <span className="text-sm text-neutral-600">
              이번주 예약 <span className="font-bold text-indigo-black">{stats?.weekVisits || 0}</span>건
            </span>
          </div>
        )}

        {/* 오늘 일정 타임라인 */}
        <Card>
          <CardHeader>
            <CardTitle>오늘 일정</CardTitle>
          </CardHeader>
          {statsLoading ? (
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

        {/* 미확인 상담 */}
        {urgentLoading ? (
          <Skeleton className="h-40" />
        ) : (
          <UrgentList
            title="미확인 상담"
            items={(urgentData?.consultations || []).map((c) => ({
              id: c.id,
              label: c.name,
              sublabel: `${CONSULTATION_TYPE_LABEL[c.consultation_type] || ''} · ${formatRelative(c.received_at)}`,
              badge: '대기중',
              badgeColor: 'bg-warning-soft text-warning',
            }))}
            onItemClick={(id) => router.push(`/consultations/${id}`)}
            emptyMessage="미확인 상담 없음"
          />
        )}

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
