'use client';

import { Topbar } from '@/components/layout/topbar';
import { HubCategoryCard } from '@/components/dashboard/hub-category-card';
import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { useHubStats, useOutstandingAlert } from '@/hooks/use-dashboard-stats';
import { Badge } from '@/components/ui/badge';
import { ShoppingCart, MessageSquare, Wrench, AlertTriangle, ClipboardList, ArrowRight } from 'lucide-react';
import Link from 'next/link';

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

export default function DashboardPage() {
  const { data: stats, isLoading } = useHubStats();
  const { data: outstanding } = useOutstandingAlert();

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

        {/* 오늘 할 일 — 대응 필요한 액션만 표시 */}
        {stats && (() => {
          const actions: Array<{ label: string; count: number; href: string; color: string }> = [];
          if (stats.orders.payDone > 0) actions.push({ label: '주문 결제 확인 → 배송 준비', count: stats.orders.payDone, href: '/orders/dashboard', color: 'bg-blue-50 text-blue-700' });
          if (stats.consultations.newIntake > 0) actions.push({ label: '신규 상담 접수 확인', count: stats.consultations.newIntake, href: '/consultations/dashboard', color: 'bg-yellow-50 text-yellow-700' });
          if (stats.consultations.needAction > 0) actions.push({ label: '상담 대응 필요', count: stats.consultations.needAction, href: '/consultations/dashboard', color: 'bg-red-50 text-red-700' });
          if (stats.repairs.intakeNew > 0) actions.push({ label: '복원수리 신규 접수 확인', count: stats.repairs.intakeNew, href: '/repairs/dashboard', color: 'bg-blue-50 text-blue-700' });
          if (stats.orders.preparing > 0) actions.push({ label: '주문 준비 → 송장 생성', count: stats.orders.preparing, href: '/orders/dashboard', color: 'bg-orange-50 text-orange-700' });
          if (actions.length === 0) return null;
          return (
            <Card>
              <div className="flex items-center gap-2 mb-3">
                <ClipboardList size={16} className="text-indigo-black" />
                <span className="text-sm font-bold">오늘 할 일</span>
                <Badge className="bg-neutral-100 text-neutral-600">{actions.length}</Badge>
              </div>
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
            </Card>
          );
        })()}

        {/* 미수금 경고 */}
        {outstanding && outstanding.length > 0 && (
          <Card>
            <div className="flex items-center gap-2 mb-3">
              <AlertTriangle size={16} className="text-amber-500" />
              <span className="text-sm font-bold text-amber-700">미수금 알림 ({outstanding.length}건)</span>
            </div>
            <div className="space-y-2">
              {outstanding.map((c) => (
                <Link
                  key={c.id}
                  href={`/customers/${c.id}`}
                  className="flex items-center justify-between p-2 rounded-lg hover:bg-amber-50 transition"
                >
                  <div>
                    <span className="text-sm font-medium">{c.name}</span>
                    {c.phone && <span className="text-xs text-neutral-400 ml-2">{c.phone}</span>}
                  </div>
                  <span className="text-sm font-bold text-amber-700">{fmtKRW(c.outstanding_balance)}</span>
                </Link>
              ))}
            </div>
          </Card>
        )}
      </div>
    </>
  );
}
