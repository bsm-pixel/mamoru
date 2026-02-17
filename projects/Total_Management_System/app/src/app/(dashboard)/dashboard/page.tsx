'use client';

import { useRouter } from 'next/navigation';
import { useQuery } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { StatsCard } from '@/components/dashboard/stats-card';
import { Skeleton } from '@/components/ui/skeleton';
import { useOrderSync } from '@/hooks/use-orders';
import { createClient } from '@/lib/supabase/client';
import { formatKRW, formatRelative, ORDER_STATUS_LABEL, ORDER_STATUS_COLOR } from '@/lib/utils/format';
import { ShoppingCart, Truck, CreditCard, Clock, RefreshCw, ArrowRight, MessageSquare, UserCheck } from 'lucide-react';
import type { Order, Consultation } from '@/lib/supabase/types';
import { CONSULTATION_STATUS_LABEL, CONSULTATION_STATUS_COLOR } from '@/lib/utils/format';

export default function DashboardPage() {
  const router = useRouter();
  const sync = useOrderSync();
  const supabase = createClient();

  // 대시보드 통계
  const { data: stats, isLoading: statsLoading } = useQuery({
    queryKey: ['dashboard-stats'],
    queryFn: async () => {
      const [totalRes, payDoneRes, shippingRes, todayRes] = await Promise.all([
        supabase.from('orders').select('*', { count: 'exact', head: true }),
        supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'pay_done'),
        supabase.from('orders').select('*', { count: 'exact', head: true }).eq('status', 'shipping'),
        supabase
          .from('orders')
          .select('*', { count: 'exact', head: true })
          .gte('ordered_at', new Date(new Date().setHours(0, 0, 0, 0)).toISOString()),
      ]);

      return {
        total: totalRes.count || 0,
        payDone: payDoneRes.count || 0,
        shipping: shippingRes.count || 0,
        today: todayRes.count || 0,
      };
    },
  });

  // 상담 통계
  const { data: consultStats } = useQuery({
    queryKey: ['dashboard-consult-stats'],
    queryFn: async () => {
      const [pendingRes, todayRes] = await Promise.all([
        supabase.from('consultations').select('*', { count: 'exact', head: true }).eq('status', 'pending_admin'),
        supabase
          .from('consultations')
          .select('*', { count: 'exact', head: true })
          .gte('received_at', new Date(new Date().setHours(0, 0, 0, 0)).toISOString()),
      ]);
      return {
        pending: pendingRes.count || 0,
        today: todayRes.count || 0,
      };
    },
  });

  // 최근 주문
  const { data: recentOrders, isLoading: recentLoading } = useQuery({
    queryKey: ['recent-orders'],
    queryFn: async () => {
      const { data } = await supabase
        .from('orders')
        .select('*')
        .order('ordered_at', { ascending: false })
        .limit(5);
      return (data || []) as Order[];
    },
  });

  // 최근 상담
  const { data: recentConsults, isLoading: consultsLoading } = useQuery({
    queryKey: ['recent-consultations'],
    queryFn: async () => {
      const { data } = await supabase
        .from('consultations')
        .select('*')
        .order('received_at', { ascending: false })
        .limit(5);
      return (data || []) as Consultation[];
    },
  });

  return (
    <>
      <Topbar title="대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-6">
        {/* 빠른 동기화 */}
        <div className="flex items-center justify-between">
          <h3 className="text-sm font-semibold text-neutral-700">오늘의 현황</h3>
          <Button
            variant="secondary"
            size="sm"
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
          >
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            동기화
          </Button>
        </div>

        {/* 통계 카드 */}
        {statsLoading ? (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {Array.from({ length: 4 }).map((_, i) => (
              <Skeleton key={i} className="h-20" />
            ))}
          </div>
        ) : (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <StatsCard
              label="전체 주문"
              value={stats?.total || 0}
              icon={ShoppingCart}
            />
            <StatsCard
              label="결제완료 (준비 필요)"
              value={stats?.payDone || 0}
              icon={CreditCard}
              color="text-info"
              bgColor="bg-info/10"
            />
            <StatsCard
              label="배송중"
              value={stats?.shipping || 0}
              icon={Truck}
              color="text-success"
              bgColor="bg-success/10"
            />
            <StatsCard
              label="오늘 주문"
              value={stats?.today || 0}
              icon={Clock}
              color="text-warning"
              bgColor="bg-warning/10"
            />
          </div>
        )}

        {/* 상담 통계 */}
        <div className="grid grid-cols-2 gap-3">
          <StatsCard
            label="대기 상담"
            value={consultStats?.pending || 0}
            icon={MessageSquare}
            color="text-warning"
            bgColor="bg-warning/10"
          />
          <StatsCard
            label="오늘 접수"
            value={consultStats?.today || 0}
            icon={UserCheck}
            color="text-info"
            bgColor="bg-info/10"
          />
        </div>

        {/* 최근 상담 */}
        <Card>
          <CardHeader>
            <CardTitle>최근 상담</CardTitle>
            <Button
              variant="ghost"
              size="sm"
              onClick={() => router.push('/consultations')}
            >
              전체보기
              <ArrowRight size={14} />
            </Button>
          </CardHeader>

          {consultsLoading ? (
            <div className="space-y-3">
              {Array.from({ length: 3 }).map((_, i) => (
                <Skeleton key={i} className="h-14" />
              ))}
            </div>
          ) : !recentConsults || recentConsults.length === 0 ? (
            <div className="text-center py-8">
              <p className="text-sm text-neutral-400">상담 데이터가 없습니다</p>
            </div>
          ) : (
            <div className="divide-y divide-neutral-100 -mx-5">
              {recentConsults.map((c) => (
                <div
                  key={c.id}
                  onClick={() => router.push(`/consultations/${c.id}`)}
                  className="flex items-center justify-between px-5 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
                >
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-semibold">{c.name}</span>
                      <Badge className={CONSULTATION_STATUS_COLOR[c.status] || ''}>
                        {CONSULTATION_STATUS_LABEL[c.status]}
                      </Badge>
                    </div>
                    <p className="text-xs text-neutral-500 mt-0.5">
                      {formatRelative(c.received_at)}
                    </p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </Card>

        {/* 최근 주문 */}
        <Card>
          <CardHeader>
            <CardTitle>최근 주문</CardTitle>
            <Button
              variant="ghost"
              size="sm"
              onClick={() => router.push('/orders')}
            >
              전체보기
              <ArrowRight size={14} />
            </Button>
          </CardHeader>

          {recentLoading ? (
            <div className="space-y-3">
              {Array.from({ length: 3 }).map((_, i) => (
                <Skeleton key={i} className="h-14" />
              ))}
            </div>
          ) : recentOrders?.length === 0 ? (
            <div className="text-center py-8">
              <p className="text-sm text-neutral-400 mb-3">
                주문 데이터가 없습니다
              </p>
              <Button
                size="sm"
                onClick={() => sync.mutate()}
                disabled={sync.isPending}
              >
                아임웹에서 동기화
              </Button>
            </div>
          ) : (
            <div className="divide-y divide-neutral-100 -mx-5">
              {recentOrders?.map((order) => (
                <div
                  key={order.id}
                  onClick={() => router.push(`/orders/${order.id}`)}
                  className="flex items-center justify-between px-5 py-3 cursor-pointer hover:bg-warm-ivory/60 transition"
                >
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-semibold">{order.orderer_name}</span>
                      <Badge className={ORDER_STATUS_COLOR[order.status] || ''}>
                        {ORDER_STATUS_LABEL[order.status]}
                      </Badge>
                    </div>
                    <p className="text-xs text-neutral-500 mt-0.5">
                      {formatRelative(order.ordered_at)}
                    </p>
                  </div>
                  <span className="text-sm font-bold">{formatKRW(order.paid_amount)}</span>
                </div>
              ))}
            </div>
          )}
        </Card>

        {/* 빠른 액션 */}
        <div className="grid grid-cols-2 gap-3">
          <Button
            variant="secondary"
            className="h-14 flex-col gap-1"
            onClick={() => router.push('/orders?status=pay_done')}
          >
            <CreditCard size={18} className="text-terracotta" />
            <span className="text-xs">결제완료 주문 처리</span>
          </Button>
          <Button
            variant="secondary"
            className="h-14 flex-col gap-1"
            onClick={() => router.push('/settings')}
          >
            <RefreshCw size={18} className="text-terracotta" />
            <span className="text-xs">동기화 설정</span>
          </Button>
        </div>
      </div>
    </>
  );
}
