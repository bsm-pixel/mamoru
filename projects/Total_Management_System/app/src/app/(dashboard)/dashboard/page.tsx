'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { HubCategoryCard } from '@/components/dashboard/hub-category-card';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { useHubStats, useOutstandingAlert, useTodayConsultations, useLowStockAlert, usePurchasingAlert, useSuppliesAlert } from '@/hooks/use-dashboard-stats';
import { useSetting, useUpdateSettings } from '@/hooks/use-settings';
import { formatPhone } from '@/lib/utils/format';
import {
  ShoppingCart, MessageSquare, Wrench, Store,
  AlertTriangle, ClipboardList, ArrowRight, CheckCircle2,
  Calendar, PackageX, Truck, PackageOpen,
} from 'lucide-react';
import Link from 'next/link';

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

export default function DashboardPage() {
  const { data: stats, isLoading } = useHubStats();

  // KPI: 월 목표 (DB 설정)
  const monthGoal = useSetting<number>('dashboard.monthly_goal', 0);
  const kpiGreen = useSetting<number>('dashboard.kpi_green', 80);
  const kpiYellow = useSetting<number>('dashboard.kpi_yellow', 50);
  const cardVisibility = useSetting<Record<string, boolean>>('dashboard.card_visibility', { sales: true, repairs: true, orders: true, consultations: true });
  const cardOrder = useSetting<string[]>('dashboard.card_order', ['orders', 'consultations', 'repairs', 'sales']);
  const updateSettings = useUpdateSettings();
  const [editingGoal, setEditingGoal] = useState(false);
  const [goalInput, setGoalInput] = useState('');
  const saveGoal = () => {
    const v = parseInt(goalInput) || 0;
    updateSettings.mutate([{ key: 'dashboard.monthly_goal', value: v }]);
    setEditingGoal(false);
  };
  const { data: outstanding } = useOutstandingAlert();
  const { data: todayConsults } = useTodayConsultations();
  const { data: lowStock } = useLowStockAlert();
  const { data: purchasingAlert } = usePurchasingAlert();
  const { data: suppliesAlert } = useSuppliesAlert();

  // 오늘 할 일 액션 생성
  const actions: Array<{ label: string; count: number; href: string; color: string }> = [];
  if (stats) {
    if (stats.orders.payDone > 0) actions.push({ label: '주문 결제 확인 → 배송 준비', count: stats.orders.payDone, href: '/orders/dashboard', color: 'bg-blue-50 text-blue-700' });
    if (stats.orders.preparing > 0) actions.push({ label: '주문 준비 → 송장 생성', count: stats.orders.preparing, href: '/orders/dashboard', color: 'bg-orange-50 text-orange-700' });
    if (stats.consultations.newIntake > 0) actions.push({ label: '신규 상담 접수 확인', count: stats.consultations.newIntake, href: '/consultations', color: 'bg-yellow-50 text-yellow-700' });
    if (stats.consultations.needAction > 0) actions.push({ label: '상담 일정 재요청', count: stats.consultations.needAction, href: '/consultations', color: 'bg-red-50 text-red-700' });
    if (stats.repairs.intakeNew > 0) actions.push({ label: '복원수리 신규 접수 확인', count: stats.repairs.intakeNew, href: '/repairs/dashboard', color: 'bg-blue-50 text-blue-700' });
    if (stats.repairs.readyToShip > 0) actions.push({ label: '복원수리 출고 대기', count: stats.repairs.readyToShip, href: '/repairs', color: 'bg-green-50 text-green-700' });
  }
  if (purchasingAlert && purchasingAlert.length > 0) {
    actions.push({ label: '매입 입고 대기', count: purchasingAlert.length, href: '/purchasing', color: 'bg-indigo-50 text-indigo-700' });
  }
  if (suppliesAlert && suppliesAlert.length > 0) {
    actions.push({ label: '부자재 주문 필요', count: suppliesAlert.length, href: '/supplies', color: 'bg-neutral-100 text-neutral-700' });
  }

  const orderSummary = stats
    ? `이번주 ${fmtKRW(stats.orders.weekAmount)} · 이번달 ${fmtKRW(stats.orders.monthAmount)}`
    : '';
  const repairSummary = stats
    ? `이번달 복원수리 매출 ${fmtKRW(stats.repairs.monthRepairAmount)} (${stats.repairs.monthRepairCount}건)`
    : '';

  return (
    <>
      <Topbar title="대시보드" />

      <div className="px-4 md:px-6 py-4 space-y-5">
        {/* KPI: 이번달 매출 달성률 */}
        {monthGoal > 0 && stats && (() => {
          const current = (stats.sales.monthAmount || 0) + (stats.repairs?.monthRepairAmount || 0);
          const pct = Math.min(Math.round((current / monthGoal) * 100), 100);
          const color = pct >= kpiGreen ? 'bg-green-500' : pct >= kpiYellow ? 'bg-yellow-500' : 'bg-red-500';
          const textColor = pct >= kpiGreen ? 'text-green-600' : pct >= kpiYellow ? 'text-yellow-600' : 'text-red-500';
          return (
            <div className="bg-white rounded-lg border border-neutral-200 p-4">
              <div className="flex items-center justify-between mb-2">
                <p className="text-xs text-neutral-500">이번달 매출 목표</p>
                <button onClick={() => { setGoalInput(String(monthGoal)); setEditingGoal(true); }} className="text-[10px] text-neutral-400 hover:text-neutral-600">수정</button>
              </div>
              <div className="flex items-end gap-3 mb-2">
                <span className={`text-2xl font-bold ${textColor}`}>{pct}%</span>
                <span className="text-sm text-neutral-500">{fmtKRW(current)} / {fmtKRW(monthGoal)}</span>
              </div>
              <div className="w-full h-2.5 bg-neutral-100 rounded-full overflow-hidden">
                <div className={`h-full rounded-full transition-all ${color}`} style={{ width: `${pct}%` }} />
              </div>
            </div>
          );
        })()}

        {/* 목표 미설정 or 편집 */}
        {(monthGoal === 0 || editingGoal) && (
          <div className="bg-neutral-50 rounded-lg border border-neutral-200 p-3 flex items-center gap-3">
            <p className="text-xs text-neutral-500 shrink-0">{editingGoal ? '월 목표 수정' : '월 매출 목표 설정'}</p>
            <input
              type="number"
              value={goalInput}
              onChange={(e) => setGoalInput(e.target.value)}
              placeholder="예: 5000000"
              className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            />
            <button onClick={saveGoal} className="px-3 py-1.5 rounded-lg bg-neutral-900 text-white text-xs font-medium">저장</button>
            {editingGoal && <button onClick={() => setEditingGoal(false)} className="text-xs text-neutral-400">취소</button>}
          </div>
        )}

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

                {/* 오늘 상담 일정 */}
                {todayConsults && todayConsults.length > 0 && (
                  <Card>
                    <div className="flex items-center gap-2 mb-3">
                      <Calendar size={16} className="text-blue-500" />
                      <span className="text-sm font-bold text-blue-700">오늘 상담 ({todayConsults.length}건)</span>
                    </div>
                    <div className="space-y-1.5">
                      {todayConsults.map((c) => {
                        const typeLabel = c.consultation_type === 'store_visit' ? '매장' : c.consultation_type === 'field_request' ? '출장' : '톡';
                        return (
                          <Link
                            key={c.id}
                            href={`/consultations/${c.id}`}
                            className="flex items-center justify-between p-2 rounded-lg hover:bg-blue-50 transition"
                          >
                            <div className="flex items-center gap-2">
                              <Badge className="bg-blue-100 text-blue-700 text-[10px]">{typeLabel}</Badge>
                              <span className="text-sm font-medium">{c.name}</span>
                            </div>
                            <span className="text-xs text-neutral-500">{c.visit_time || '-'}</span>
                          </Link>
                        );
                      })}
                    </div>
                  </Card>
                )}
              </div>

              {/* 우측: 현황 카드 — 설정 기반 순서/표시 */}
              <div className="w-full lg:w-[520px] shrink-0 space-y-3">
                <h3 className="text-xs font-semibold text-neutral-400 uppercase tracking-wider">현황 요약</h3>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                {(() => {
                  const CARD_MAP: Record<string, React.ReactNode> = {
                    orders: (
                      <HubCategoryCard key="orders" title="주문" icon={ShoppingCart} href="/orders/dashboard"
                        stats={[
                          { label: '결제완료', value: stats?.orders.payDone || 0, color: 'text-info' },
                          { label: '준비중', value: stats?.orders.preparing || 0, color: 'text-warning' },
                          { label: '배송중', value: stats?.orders.shipping || 0, color: 'text-terracotta' },
                          { label: '완료', value: stats?.orders.delivered || 0, color: 'text-success' },
                        ]}
                        summary={orderSummary}
                      />
                    ),
                    consultations: (
                      <HubCategoryCard key="consultations" title="상담" icon={MessageSquare} href="/consultations"
                        stats={[
                          { label: '신규접수', value: stats?.consultations.newIntake || 0, color: 'text-warning' },
                          { label: '상담예정', value: stats?.consultations.confirmed || 0, color: 'text-info' },
                          { label: '일정재요청', value: stats?.consultations.needAction || 0, color: 'text-error' },
                        ]}
                      />
                    ),
                    repairs: (
                      <HubCategoryCard key="repairs" title="복원수리" icon={Wrench} href="/repairs/dashboard"
                        stats={[
                          { label: '신규접수', value: stats?.repairs.intakeNew || 0, color: 'text-info' },
                          { label: '진행중', value: stats?.repairs.workingCount || 0, color: 'text-terracotta' },
                          { label: '출고대기', value: stats?.repairs.readyToShip || 0, color: 'text-warning' },
                        ]}
                        summary={repairSummary}
                      />
                    ),
                    sales: (
                      <HubCategoryCard key="sales" title="오프라인 판매" icon={Store} href="/sales"
                        stats={[
                          { label: '이번달', value: stats?.sales.monthCount || 0, color: 'text-indigo-black' },
                        ]}
                        summary={stats ? `이번달 ${fmtKRW(stats.sales.monthAmount)}` : ''}
                      />
                    ),
                  };
                  const defaultOrder = ['orders', 'consultations', 'repairs', 'sales'];
                  const order = cardOrder.length > 0 ? cardOrder : defaultOrder;
                  return order
                    .filter((key) => cardVisibility[key] !== false)
                    .map((key) => CARD_MAP[key])
                    .filter(Boolean);
                })()}
                </div>
                {/* 저재고 알림 */}
                {lowStock && lowStock.length > 0 && (
                  <Card>
                    <div className="flex items-center justify-between mb-3">
                      <div className="flex items-center gap-2">
                        <PackageX size={16} className="text-red-500" />
                        <span className="text-sm font-bold text-red-700">저재고 ({lowStock.length})</span>
                      </div>
                      <Link href="/purchasing/new" className="text-xs text-terracotta hover:underline">발주 작성 →</Link>
                    </div>
                    <div className="space-y-1.5">
                      {lowStock.map((p) => (
                        <Link
                          key={p.id}
                          href={`/products/${p.id}`}
                          className="flex items-center justify-between p-2 rounded-lg hover:bg-red-50 transition"
                        >
                          <div>
                            <span className="text-sm font-medium">{p.name}</span>
                            <span className="text-xs text-neutral-400 ml-2">{p.sku}</span>
                          </div>
                          <Badge className={p.stock_quantity === 0 ? 'bg-red-100 text-red-700' : 'bg-amber-100 text-amber-700'}>
                            {p.stock_quantity}개
                          </Badge>
                        </Link>
                      ))}
                    </div>
                  </Card>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
}
