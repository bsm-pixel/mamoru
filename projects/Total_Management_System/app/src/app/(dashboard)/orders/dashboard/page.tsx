'use client';

import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { StatsCard } from '@/components/dashboard/stats-card';
import { PipelineBar } from '@/components/dashboard/pipeline-bar';
import { UrgentList } from '@/components/dashboard/urgent-list';
import { Skeleton } from '@/components/ui/skeleton';
import { Button } from '@/components/ui/button';
import { useOrderDashboardStats } from '@/hooks/use-dashboard-stats';
import { useOrders } from '@/hooks/use-orders';
import { formatRelative, formatKRW } from '@/lib/utils/format';
import { CreditCard, Package, Truck, CheckCircle, ArrowRight, ShoppingCart } from 'lucide-react';

export default function OrderDashboardPage() {
  const router = useRouter();
  const { data: stats, isLoading: statsLoading } = useOrderDashboardStats();
  const { data: urgentData, isLoading: urgentLoading } = useOrders({ status: 'pay_done', limit: 5 });

  return (
    <>
      <Topbar title="주문 대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-6">
        {/* 파이프라인 */}
        <Card>
          <CardHeader>
            <CardTitle>주문 처리 현황</CardTitle>
          </CardHeader>
          {statsLoading ? (
            <Skeleton className="h-16" />
          ) : (
            <PipelineBar
              stages={(stats?.pipeline || []).map((s) => ({ ...s, href: `/orders?status=${s.status}` }))}
            />
          )}
        </Card>

        {/* 통계 카드 */}
        {statsLoading ? (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {Array.from({ length: 4 }).map((_, i) => (
              <Skeleton key={i} className="h-20" />
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <StatsCard label="결제완료" value={stats?.payDone || 0} icon={CreditCard} color="text-info" bgColor="bg-info/10" />
            <StatsCard label="준비중" value={stats?.preparing || 0} icon={Package} color="text-warning" bgColor="bg-warning/10" />
            <StatsCard label="배송중" value={stats?.shipping || 0} icon={Truck} color="text-terracotta" bgColor="bg-terracotta/10" />
            <StatsCard label="오늘 주문" value={stats?.todayOrders || 0} icon={ShoppingCart} color="text-success" bgColor="bg-success/10" />
          </div>
        )}

        {/* 처리 필요 주문 */}
        {urgentLoading ? (
          <Skeleton className="h-40" />
        ) : (
          <UrgentList
            title="결제완료 — 처리 필요"
            items={(urgentData?.orders || []).map((o) => ({
              id: o.id,
              label: o.orderer_name || '주문자',
              sublabel: `${formatKRW(o.paid_amount)} · ${formatRelative(o.ordered_at)}`,
              badge: '결제완료',
              badgeColor: 'bg-info-soft text-info',
            }))}
            onItemClick={(id) => router.push(`/orders/${id}`)}
            emptyMessage="처리 대기 주문 없음"
          />
        )}

        {/* 전체 목록 링크 */}
        <Button
          variant="ghost"
          className="w-full justify-center gap-1"
          onClick={() => router.push('/orders')}
        >
          전체 주문 목록
          <ArrowRight size={14} />
        </Button>
      </div>
    </>
  );
}
