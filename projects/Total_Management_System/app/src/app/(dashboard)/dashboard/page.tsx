'use client';

import { Topbar } from '@/components/layout/topbar';
import { HubCategoryCard } from '@/components/dashboard/hub-category-card';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { useHubStats, useOutstandingAlert } from '@/hooks/use-dashboard-stats';
import { formatPhone } from '@/lib/utils/format';
import {
  ShoppingCart, MessageSquare, Wrench, Store,
  AlertTriangle, ClipboardList, ArrowRight, CheckCircle2,
} from 'lucide-react';
import Link from 'next/link';

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

export default function DashboardPage() {
  const { data: stats, isLoading } = useHubStats();
  const { data: outstanding } = useOutstandingAlert();

  // 오늘 할 일 액션 생성
  const actions: Array<{ label: string; count: number; href: string; color: string }> = [];
  if (stats) {
    if (stats.orders.payDone > 0) actions.push({ label: '주문 결제 확인 → 배송 준비', count: stats.orders.payDone, href: '/orders/dashboard', color: 'bg-blue-50 text-blue-700' });
    if (stats.orders.preparing > 0) actions.push({ label: '주문 준비 → 송장 생성', count: stats.orders.preparing, href: '/orders/dashboard', color: 'bg-orange-50 text-orange-700' });
    if (stats.consultations.newIntake > 0) actions.push({ label: '신규 상담 접수 확인', count: stats.consultations.newIntake, href: '/consultations/dashboard', color: 'bg-yellow-50 text-yellow-700' });
    if (stats.consultations.needAction > 0) actions.push({ label: '상담 대응 필요', count: stats.consultations.needAction, href: '/consultations/dashboard', color: 'bg-red-50 text-red-700' });
    if (stats.repairs.intakeNew > 0) actions.push({ label: '복원수리 신규 접수 확인', count: stats.repairs.intakeNew, href: '/repairs/dashboard', color: 'bg-blue-50 text-blue-700' });
    if (stats.repairs.readyToShip > 0) actions.push({ label: '복원수리 출고 대기', count: stats.repairs.readyToShip, href: '/repairs', color: 'bg-green-50 text-green-700' });
  }

  const orderSummary = stats
    ? `이번주 ${fmtKRW(stats.orders.weekAmount)} · 이번달 ${fmtKRW(stats.orders.monthAmount)}`
    : '';
  const repairSummary = stats
    ? `이번주 복원 ${stats.repairs.weekRepairTotal}건 (마모루 ${stats.repairs.weekRepairMamoru} / 타사 ${stats.repairs.weekRepairOther})`
    : '';

  return (
    <>
      <Topbar title="대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-5">
        {isLoading ? (
          <div className="space-y-4">
            <Skeleton className="h-40 w-full" />
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
              <Skeleton className="h-32" />
              <Skeleton className="h-32" />
            </div>
          </div>
        ) : (
          <>
            {/* PC: 2단 레이아웃 / 모바일: 수직 */}
            <div className="flex flex-col lg:flex-row gap-5">

              {/* 좌측: 오늘 할 일 + 미수금 */}
              <div className="flex-1 min-w-0 space-y-4">
                {/* 오늘 할 일 */}
                <Card>
                  <div className="flex items-center gap-2 mb-3">
                    <ClipboardList size={18} className="text-indigo-black" />
                    <h3 className="text-sm font-bold">오늘 할 일</h3>
                    {actions.length > 0 && (
                      <Badge className="bg-red-100 text-red-700">{actions.length}</Badge>
                    )}
                  </div>
                  {actions.length > 0 ? (
                    <div className="space-y-1.5">
                      {actions.map((a) => (
                        <Link
                          key={a.label}
                          href={a.href}
                          className={`flex items-center justify-between p-2.5 rounded-lg ${a.color} transition hover:opacity-80`}
                        >
                          <div className="flex items-center gap-2">
                            <span className="text-sm font-medium">{a.label}</span>
                            <Badge className="bg-white/60 text-inherit">{a.count}건</Badge>
                          </div>
                          <ArrowRight size={14} className="opacity-40" />
                        </Link>
                      ))}
                    </div>
                  ) : (
                    <div className="flex items-center gap-2 py-4 justify-center text-green-600">
                      <CheckCircle2 size={18} />
                      <span className="text-sm font-medium">모든 업무가 처리되었습니다</span>
                    </div>
                  )}
                </Card>

                {/* 미수금 경고 */}
                {outstanding && outstanding.length > 0 && (
                  <Card>
                    <div className="flex items-center gap-2 mb-3">
                      <AlertTriangle size={16} className="text-amber-500" />
                      <span className="text-sm font-bold text-amber-700">미수금 ({outstanding.length}건)</span>
                    </div>
                    <div className="space-y-1.5">
                      {outstanding.map((c) => (
                        <Link
                          key={c.id}
                          href={`/customers/${c.id}`}
                          className="flex items-center justify-between p-2 rounded-lg hover:bg-amber-50 transition"
                        >
                          <div>
                            <span className="text-sm font-medium">{c.name}</span>
                            {c.phone && <span className="text-xs text-neutral-400 ml-2">{formatPhone(c.phone)}</span>}
                          </div>
                          <span className="text-sm font-bold text-amber-700">{fmtKRW(c.outstanding_balance)}</span>
                        </Link>
                      ))}
                    </div>
                  </Card>
                )}
              </div>

              {/* 우측: 현황 카드 세로 */}
              <div className="w-full lg:w-96 shrink-0 space-y-3">
                <h3 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">현황 요약</h3>

                <HubCategoryCard
                  title="주문"
                  icon={ShoppingCart}
                  href="/orders/dashboard"
                  stats={[
                    { label: '결제완료', value: stats?.orders.payDone || 0, color: 'text-info' },
                    { label: '준비중', value: stats?.orders.preparing || 0, color: 'text-warning' },
                    { label: '배송중', value: stats?.orders.shipping || 0, color: 'text-terracotta' },
                    { label: '완료', value: stats?.orders.delivered || 0, color: 'text-success' },
                  ]}
                  summary={orderSummary}
                />

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

                <HubCategoryCard
                  title="복원수리"
                  icon={Wrench}
                  href="/repairs/dashboard"
                  stats={[
                    { label: '신규접수', value: stats?.repairs.intakeNew || 0, color: 'text-info' },
                    { label: '진행중', value: stats?.repairs.workingCount || 0, color: 'text-terracotta' },
                    { label: '출고대기', value: stats?.repairs.readyToShip || 0, color: 'text-warning' },
                  ]}
                  summary={repairSummary}
                />

                <HubCategoryCard
                  title="오프라인 판매"
                  icon={Store}
                  href="/sales"
                  stats={[
                    { label: '이번달', value: stats?.sales.monthCount || 0, color: 'text-indigo-black' },
                  ]}
                  summary={stats ? `이번달 ${fmtKRW(stats.sales.monthAmount)}` : ''}
                />
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
}
