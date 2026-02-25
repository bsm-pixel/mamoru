'use client';

import { Topbar } from '@/components/layout/topbar';
import { HubCategoryCard } from '@/components/dashboard/hub-category-card';
import { Skeleton } from '@/components/ui/skeleton';
import { useHubStats } from '@/hooks/use-dashboard-stats';
import { ShoppingCart, MessageSquare, Wrench } from 'lucide-react';

export default function DashboardPage() {
  const { data: stats, isLoading } = useHubStats();

  return (
    <>
      <Topbar title="대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <h3 className="text-sm font-semibold text-neutral-700">오늘의 현황</h3>

        {isLoading ? (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            {Array.from({ length: 3 }).map((_, i) => (
              <Skeleton key={i} className="h-28" />
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            <HubCategoryCard
              title="주문"
              icon={ShoppingCart}
              href="/orders/dashboard"
              stats={[
                { label: '결제완료', value: stats?.orders.payDone || 0, color: 'text-info' },
                { label: '배송중', value: stats?.orders.shipping || 0, color: 'text-success' },
              ]}
            />
            <HubCategoryCard
              title="상담"
              icon={MessageSquare}
              href="/consultations/dashboard"
              stats={[
                { label: '미확인', value: stats?.consultations.pendingAdmin || 0, color: 'text-warning' },
                { label: '오늘 예약', value: stats?.consultations.todayVisit || 0, color: 'text-info' },
              ]}
            />
            <HubCategoryCard
              title="복원수리"
              icon={Wrench}
              href="/repairs/dashboard"
              stats={[
                { label: '접수처리', value: stats?.repairs.intake || 0, color: 'text-info' },
                { label: '비용안내 대기', value: stats?.repairs.costNotified || 0, color: 'text-warning' },
              ]}
            />
          </div>
        )}
      </div>
    </>
  );
}
