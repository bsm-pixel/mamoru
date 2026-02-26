'use client';

import { Topbar } from '@/components/layout/topbar';
import { HubCategoryCard } from '@/components/dashboard/hub-category-card';
import { Skeleton } from '@/components/ui/skeleton';
import { useHubStats } from '@/hooks/use-dashboard-stats';
import { ShoppingCart, MessageSquare, Wrench } from 'lucide-react';

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

export default function DashboardPage() {
  const { data: stats, isLoading } = useHubStats();

  // R3: 주문 요약
  const orderSummary = stats
    ? `이번주 ${fmtKRW(stats.orders.weekAmount)} · 이번달 ${fmtKRW(stats.orders.monthAmount)}`
    : '';

  // R3: 복원수리 요약
  const repairSummary = stats
    ? `이번주 복원 ${stats.repairs.weekRepairTotal}건 (마모루 ${stats.repairs.weekRepairMamoru} / 타사 ${stats.repairs.weekRepairOther})`
    : '';

  return (
    <>
      <Topbar title="대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <h3 className="text-sm font-semibold text-neutral-700">오늘의 현황</h3>

        {isLoading ? (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            {Array.from({ length: 3 }).map((_, i) => (
              <Skeleton key={i} className="h-32" />
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
            {/* R3: 판매 카드 — 4단계 + 주/월 금액 */}
            <HubCategoryCard
              title="판매"
              icon={ShoppingCart}
              href="/orders/dashboard"
              stats={[
                { label: '주문결제', value: stats?.orders.payDone || 0, color: 'text-info' },
                { label: '준비중', value: stats?.orders.preparing || 0, color: 'text-warning' },
                { label: '배송중', value: stats?.orders.shipping || 0, color: 'text-terracotta' },
                { label: '배송완료', value: stats?.orders.delivered || 0, color: 'text-success' },
              ]}
              summary={orderSummary}
            />

            {/* R3: 상담 카드 — 신규접수 + 상담예정 + 대응필요 */}
            <HubCategoryCard
              title="상담"
              icon={MessageSquare}
              href="/consultations/dashboard"
              stats={[
                { label: '신규접수', value: stats?.consultations.newIntake || 0, color: 'text-warning' },
                { label: '상담예정', value: stats?.consultations.confirmed || 0, color: 'text-info' },
                { label: '대응필요', value: stats?.consultations.needAction || 0, color: 'text-error' },
              ]}
            />

            {/* R3: 복원수리 카드 — 신규접수 + 진행대기 + 진행중(가위수량) */}
            <HubCategoryCard
              title="복원수리"
              icon={Wrench}
              href="/repairs/dashboard"
              stats={[
                { label: '신규접수', value: stats?.repairs.intakeNew || 0, color: 'text-info' },
                { label: '진행대기', value: stats?.repairs.pendingInbound || 0, color: 'text-warning' },
                { label: `진행중`, value: stats?.repairs.workingCount || 0, color: 'text-terracotta' },
              ]}
              summary={repairSummary}
            />
          </div>
        )}
      </div>
    </>
  );
}
