'use client';

/**
 * /dashboard — TMS 메인 대시보드
 *
 * 시안 B+ (2026-05-26 사장님 채택, 시안 디자인 모니터 비교 후 확정):
 *   - 1행: 매출 KPI(SVG 도넛, 5/12) + 4카드(7/12, 카테고리별 컴팩트)
 *   - 2행: DashboardCalendarPanel (달력 + 선택일 타임라인, 기본=오늘)
 *   - 3행: 미수금 + 할일 + 알림 5종(가로 5분할)
 *
 * TMS 디자인 가이드라인 (memory/feedback_tms_design_direction.md):
 *   마모루 모노크롬(stone-50/white/stone-900) 베이스 + 절제된 상태색.
 *   회계 합산 로직(RPC 077·078·080·088) 무수정 — UI 재배치 + 컴포넌트 분리만.
 */

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import Link from 'next/link';
import {
  ShoppingCart, MessageSquare, Wrench, Store,
  CheckCircle2, AlertTriangle,
  PackageX, Truck, PackageOpen, Star, ClipboardList,
  Plus,
} from 'lucide-react';

import { Topbar } from '@/components/layout/topbar';
import { Skeleton } from '@/components/ui/skeleton';
import { StatCard } from '@/components/ui/stat-card';
import { DashboardCalendarPanel } from '@/components/dashboard/dashboard-calendar';
import {
  useHubStats, useOutstandingAlert,
  useLowStockAlert, usePurchasingAlert, useSuppliesAlert,
  useWaybillAlert, useNewReviewAlert,
} from '@/hooks/use-dashboard-stats';
import { useSetting, useUpdateSettings } from '@/hooks/use-settings';
import { formatPhone } from '@/lib/utils/format';
import { useActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';

function fmtKRW(n: number) {
  if (n >= 10000) return `₩${Math.round(n / 10000)}만`;
  return `₩${n.toLocaleString()}`;
}

/* ──────────────────────────────────────────────────────────
 * 매출 KPI 도넛 (1행 좌측)
 * ──────────────────────────────────────────────────────── */
function RevenueKPIDonut({
  current, monthGoal, b2c, b2b, repair, kpiGreen, kpiYellow, onEditClick,
}: {
  current: number; monthGoal: number;
  b2c: number; b2b: number; repair: number;
  kpiGreen: number; kpiYellow: number;
  onEditClick: () => void;
}) {
  const pct = monthGoal > 0 ? Math.min(Math.round((current / monthGoal) * 100), 100) : 0;
  const ringColor = pct >= kpiGreen ? '#10B981' : pct >= kpiYellow ? '#F59E0B' : '#EF4444';
  const textColor = pct >= kpiGreen ? 'text-emerald-600' : pct >= kpiYellow ? 'text-amber-600' : 'text-rose-500';
  const radius = 30;
  const circ = 2 * Math.PI * radius;
  const dash = (pct / 100) * circ;

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-5 h-full flex flex-col">
      <div className="flex items-center justify-between mb-3">
        <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">이번달 매출</p>
        <button onClick={onEditClick} className="text-[10px] text-stone-400 hover:text-stone-700 transition">목표 수정</button>
      </div>
      <div className="flex items-center gap-4 flex-1">
        <div className="relative w-[88px] h-[88px] shrink-0">
          <svg viewBox="0 0 72 72" className="w-full h-full -rotate-90">
            <circle cx="36" cy="36" r={radius} fill="none" stroke="#F5F5F4" strokeWidth="6" />
            <circle
              cx="36" cy="36" r={radius} fill="none"
              stroke={ringColor} strokeWidth="6" strokeLinecap="round"
              strokeDasharray={`${dash} ${circ}`}
            />
          </svg>
          <div className="absolute inset-0 flex items-center justify-center">
            <span className={`text-xl font-bold ${textColor}`}>{pct}<span className="text-xs text-stone-500">%</span></span>
          </div>
        </div>
        <div className="flex-1 min-w-0">
          <p className="text-lg font-bold text-stone-900 leading-tight">{fmtKRW(current)}</p>
          <p className="text-[11px] text-stone-500 mt-0.5">목표 {fmtKRW(monthGoal)}</p>
          <div className="grid grid-cols-3 gap-1 mt-3">
            <div className="text-center">
              <p className="text-[9px] text-stone-400">B2C</p>
              <p className="text-[11px] font-semibold text-stone-700">{fmtKRW(b2c)}</p>
            </div>
            <div className="text-center border-l border-stone-100">
              <p className="text-[9px] text-stone-400">B2B</p>
              <p className="text-[11px] font-semibold text-stone-700">{fmtKRW(b2b)}</p>
            </div>
            <div className="text-center border-l border-stone-100">
              <p className="text-[9px] text-stone-400">수리</p>
              <p className="text-[11px] font-semibold text-stone-700">{fmtKRW(repair)}</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 매출 목표 미설정/편집 박스 (도넛 자리 대체)
 * ──────────────────────────────────────────────────────── */
function GoalEditBox({
  goalInput, setGoalInput, onSave, onCancel, isEdit,
}: {
  goalInput: string; setGoalInput: (v: string) => void;
  onSave: () => void; onCancel?: () => void; isEdit: boolean;
}) {
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-5 h-full flex items-center gap-3">
      <p className="text-xs text-stone-500 shrink-0">{isEdit ? '월 목표 수정' : '월 매출 목표 설정'}</p>
      <input
        type="number"
        value={goalInput}
        onChange={(e) => setGoalInput(e.target.value)}
        placeholder="예: 5000000"
        className="flex-1 h-8 px-3 rounded-lg border border-stone-200 text-sm focus:outline-none focus:border-stone-400 transition"
      />
      <button onClick={onSave} className="px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition">저장</button>
      {isEdit && onCancel && (
        <button onClick={onCancel} className="text-xs text-stone-400 hover:text-stone-600 transition">취소</button>
      )}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 미수금 카드 (3행 좌)
 * ──────────────────────────────────────────────────────── */
function OutstandingCard({
  outstanding,
}: {
  outstanding?: Array<{ id: string; name: string; phone: string | null; outstanding_balance: number }>;
}) {
  const items = outstanding || [];
  const actTypes = useActivityTypes(items.map((c) => c.phone));
  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-1.5">
          <AlertTriangle size={13} className={items.length > 0 ? 'text-rose-500' : 'text-stone-400'} />
          <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">미수금</p>
        </div>
        {items.length > 0 ? (
          <span className="text-[10px] px-2 py-0.5 rounded-full bg-rose-50 text-rose-700 font-semibold">{items.length}건</span>
        ) : (
          <span className="text-[10px] px-2 py-0.5 rounded-full bg-stone-100 text-stone-500 font-semibold">0건</span>
        )}
      </div>
      {items.length === 0 ? (
        <div className="py-6 text-center text-xs text-stone-400">미수금이 없습니다</div>
      ) : (
        <div className="space-y-1.5">
          {items.map((c) => (
            <Link
              key={c.id}
              href={`/customers/${c.id}`}
              className="flex items-center justify-between p-2 rounded-lg hover:bg-rose-50/40 transition"
            >
              <div className="min-w-0">
                <div className="flex items-center gap-1 min-w-0">
                  <p className="text-xs font-semibold text-stone-800 truncate">{c.name}</p>
                  <ActivityChips types={actTypes(c.phone)} className="shrink-0" />
                </div>
                {c.phone && <p className="text-[10px] text-stone-400 truncate">{formatPhone(c.phone)}</p>}
              </div>
              <span className="text-xs font-bold text-rose-600 shrink-0">{fmtKRW(c.outstanding_balance)}</span>
            </Link>
          ))}
        </div>
      )}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 할 일 메모 카드 (3행 중)
 * ──────────────────────────────────────────────────────── */
function TodoCard() {
  const queryClient = useQueryClient();
  const [input, setInput] = useState('');
  const [confirmId, setConfirmId] = useState<string | null>(null);

  const { data } = useQuery({
    queryKey: ['dashboard-todos'],
    staleTime: 30_000,
    queryFn: async () => {
      const res = await fetch('/api/dashboard/todos');
      if (!res.ok) return { todos: [] };
      return res.json() as Promise<{ todos: Array<{ id: string; text: string; created_at: string }> }>;
    },
  });

  const addTodo = useMutation({
    mutationFn: async (text: string) => {
      const res = await fetch('/api/dashboard/todos', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text }),
      });
      if (!res.ok) throw new Error();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['dashboard-todos'] });
      setInput('');
    },
  });

  const deleteTodo = useMutation({
    mutationFn: async (id: string) => {
      await fetch('/api/dashboard/todos', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id }),
      });
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['dashboard-todos'] });
      setConfirmId(null);
    },
  });

  const todos = data?.todos || [];

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full flex flex-col">
      <div className="flex items-center justify-between mb-3">
        <div className="flex items-center gap-1.5">
          <CheckCircle2 size={13} className="text-stone-400" />
          <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">할 일 메모</p>
        </div>
        <span className="text-[10px] px-2 py-0.5 rounded-full bg-stone-100 text-stone-600 font-semibold">{todos.length}</span>
      </div>
      <div className="flex gap-1.5 mb-3">
        <input
          type="text"
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter' && input.trim()) addTodo.mutate(input.trim()); }}
          placeholder="할 일 입력..."
          className="flex-1 h-7 px-2.5 rounded-lg border border-stone-200 text-xs placeholder:text-stone-400 focus:outline-none focus:border-stone-400 transition"
        />
        <button
          onClick={() => { if (input.trim()) addTodo.mutate(input.trim()); }}
          disabled={!input.trim() || addTodo.isPending}
          className="px-2 h-7 rounded-lg bg-stone-900 text-white text-[10px] font-semibold hover:bg-stone-800 transition disabled:opacity-40 flex items-center gap-0.5"
        >
          <Plus size={11} />추가
        </button>
      </div>
      {todos.length === 0 ? (
        <div className="py-2 text-center text-xs text-stone-400 flex-1">등록된 할 일이 없습니다</div>
      ) : (
        <div className="space-y-1 flex-1">
          {todos.map((todo) => (
            <div key={todo.id} className="flex items-center gap-2 py-1 group">
              <button
                onClick={() => setConfirmId(todo.id)}
                className="w-3.5 h-3.5 rounded border border-stone-300 shrink-0 hover:border-emerald-500 hover:bg-emerald-50 transition flex items-center justify-center"
                aria-label="완료"
              />
              <span className="text-xs text-stone-700 flex-1">{todo.text}</span>
            </div>
          ))}
        </div>
      )}

      {/* 완료 확인 모달 */}
      {confirmId && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/30" onClick={() => setConfirmId(null)}>
          <div className="bg-white rounded-2xl shadow-xl p-5 w-72" onClick={(e) => e.stopPropagation()}>
            <p className="text-sm font-semibold mb-3 text-stone-800">이 할일을 완료했습니까?</p>
            <div className="flex gap-2">
              <button
                onClick={() => deleteTodo.mutate(confirmId)}
                className="flex-1 py-2 rounded-lg bg-emerald-600 text-white text-xs font-semibold hover:bg-emerald-700 transition"
              >완료</button>
              <button
                onClick={() => setConfirmId(null)}
                className="flex-1 py-2 rounded-lg bg-stone-100 text-stone-600 text-xs font-semibold hover:bg-stone-200 transition"
              >취소</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 알림 5종 카드 (3행 우) — 가로 5분할
 * ──────────────────────────────────────────────────────── */
function AlertsCard({
  lowStock, waybill, newReviews, purchasing, supplies,
}: {
  lowStock?: Array<unknown>;
  waybill?: { remaining: number } | null;
  newReviews?: { count: number } | null;
  purchasing?: Array<unknown>;
  supplies?: Array<unknown>;
}) {
  const lowStockCount = lowStock?.length || 0;
  const waybillCount = waybill?.remaining ?? 0;
  const reviewCount = newReviews?.count || 0;
  const purchasingCount = purchasing?.length || 0;
  const suppliesCount = supplies?.length || 0;

  const alerts: Array<{
    icon: typeof PackageX; label: string; count: number; href: string;
    active: boolean; color: 'rose' | 'amber' | 'yellow' | 'blue' | 'stone';
  }> = [
    { icon: PackageX,      label: '저재고',   count: lowStockCount,  href: '/inventory?low=1', active: lowStockCount > 0,  color: 'rose'   },
    { icon: Truck,         label: '운송장',   count: waybillCount,    href: '/settings',       active: waybillCount < 100, color: 'amber'  },
    { icon: Star,          label: '신규 후기', count: reviewCount,     href: '/reviews',        active: reviewCount > 0,    color: 'yellow' },
    { icon: PackageOpen,   label: '매입 대기', count: purchasingCount, href: '/purchasing',     active: purchasingCount > 0,color: 'blue'   },
    { icon: ClipboardList, label: '부자재',   count: suppliesCount,    href: '/supplies',       active: suppliesCount > 0,  color: 'stone'  },
  ];

  const palette = {
    rose:   { text: 'text-rose-700',   bg: 'bg-rose-50',   icon: 'text-rose-500'   },
    amber:  { text: 'text-amber-700',  bg: 'bg-amber-50',  icon: 'text-amber-500'  },
    yellow: { text: 'text-yellow-700', bg: 'bg-yellow-50', icon: 'text-yellow-500' },
    blue:   { text: 'text-blue-700',   bg: 'bg-blue-50',   icon: 'text-blue-500'   },
    stone:  { text: 'text-stone-700',  bg: 'bg-stone-100', icon: 'text-stone-500'  },
  } as const;
  const inactivePalette = { text: 'text-stone-400', bg: 'bg-stone-50', icon: 'text-stone-300' };

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full">
      <div className="flex items-center justify-between mb-3">
        <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">시스템 알림</p>
      </div>
      <div className="grid grid-cols-5 gap-1.5">
        {alerts.map((a) => {
          const c = a.active ? palette[a.color] : inactivePalette;
          return (
            <Link
              key={a.label}
              href={a.href}
              className={`${c.bg} rounded-lg p-2 hover:opacity-80 transition`}
            >
              <div className="flex items-center justify-between mb-1">
                <a.icon size={11} className={c.icon} />
                <span className={`text-xs font-bold ${c.text}`}>{a.count}</span>
              </div>
              <p className={`text-[9px] font-medium ${c.text} truncate`}>{a.label}</p>
            </Link>
          );
        })}
      </div>
    </div>
  );
}

/* ──────────────────────────────────────────────────────────
 * 페이지 본체
 * ──────────────────────────────────────────────────────── */
export default function DashboardPage() {
  const { data: stats, isLoading } = useHubStats();
  const monthGoal = useSetting<number>('dashboard.monthly_goal', 0);
  const kpiGreen = useSetting<number>('dashboard.kpi_green', 80);
  const kpiYellow = useSetting<number>('dashboard.kpi_yellow', 50);
  const cardVisibility = useSetting<Record<string, boolean>>('dashboard.card_visibility', {
    sales: true, repairs: true, orders: true, consultations: true,
  });
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
  const { data: lowStock } = useLowStockAlert();
  const { data: purchasingAlert } = usePurchasingAlert();
  const { data: suppliesAlert } = useSuppliesAlert();
  const { data: waybillAlert } = useWaybillAlert();
  const { data: newReviews } = useNewReviewAlert();

  // 매출 3분할 (회계 합산 로직 기존 그대로 유지)
  // B2C 제품 + B2B 제품(딜러/아카데미+납품) + 복원수리 전체(A 접수+B 판매RS+C 납품RS)
  const online = stats?.sales.salesOnline || 0;         // 아임웹 온라인 주문 = B2C 제품매출에 합산
  const b2c = (stats?.sales.salesB2C || 0) + online;
  const b2b = stats?.sales.salesB2B || 0;
  const repair = stats?.repairs?.monthRepairAmount || 0;
  const current = b2c + b2b + repair;

  // 4카드 정의 (설정 기반 순서/표시) — 공통 StatCard 재사용
  const CARD_DEF: Record<string, React.ReactNode> = {
    orders: (
      <StatCard
        key="orders" label="주문" icon={ShoppingCart} accent="blue" href="/orders/dashboard"
        value={stats?.orders.payDone || 0} primarySub="결제완료"
        secondarySub={stats ? `준비 ${stats.orders.preparing} · 완료 ${stats.orders.delivered}` : undefined}
      />
    ),
    consultations: (
      <StatCard
        key="consultations" label="상담" icon={MessageSquare} accent="amber" href="/consultations"
        value={stats?.consultations.newIntake || 0} primarySub="신규접수"
        secondarySub={stats ? `예정 ${stats.consultations.confirmed} · 재요청 ${stats.consultations.needAction}` : undefined}
      />
    ),
    repairs: (
      <StatCard
        key="repairs" label="복원수리" icon={Wrench} accent="emerald" href="/repairs/dashboard"
        value={stats?.repairs.readyToShip || 0} primarySub="출고대기"
        secondarySub={stats ? `이번달 ${stats.repairs.monthRepairCount}자루` : undefined}
      />
    ),
    sales: (
      <StatCard
        key="sales" label="제품 판매" icon={Store} accent="stone" href="/sales"
        value={stats?.sales.monthCount || 0} primarySub="이번달 판매"
        secondarySub={stats ? `B2C ${fmtKRW(b2c)} · B2B ${fmtKRW(b2b)}` : undefined}
      />
    ),
  };
  const defaultOrder = ['orders', 'consultations', 'repairs', 'sales'];
  const order = cardOrder.length > 0 ? cardOrder : defaultOrder;
  const visibleCards = order
    .filter((key) => cardVisibility[key] !== false)
    .map((key) => CARD_DEF[key])
    .filter(Boolean);

  return (
    <>
      <Topbar title="대시보드" />
      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4">
        {isLoading ? (
          <div className="space-y-3">
            <Skeleton className="h-40 w-full" />
            <Skeleton className="h-64 w-full" />
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
              <Skeleton className="h-40" />
              <Skeleton className="h-40" />
              <Skeleton className="h-40" />
            </div>
          </div>
        ) : (
          <div className="space-y-3">
            {/* 1행: 매출 KPI + 4카드 */}
            <div className="grid grid-cols-12 gap-3">
              <div className="col-span-12 lg:col-span-5">
                {monthGoal > 0 && !editingGoal ? (
                  <RevenueKPIDonut
                    current={current} monthGoal={monthGoal}
                    b2c={b2c} b2b={b2b} repair={repair}
                    kpiGreen={kpiGreen} kpiYellow={kpiYellow}
                    onEditClick={() => { setGoalInput(String(monthGoal)); setEditingGoal(true); }}
                  />
                ) : (
                  <GoalEditBox
                    goalInput={goalInput} setGoalInput={setGoalInput}
                    onSave={saveGoal}
                    onCancel={editingGoal ? () => setEditingGoal(false) : undefined}
                    isEdit={editingGoal}
                  />
                )}
              </div>
              <div className="col-span-12 lg:col-span-7 grid grid-cols-2 sm:grid-cols-4 gap-3">
                {visibleCards}
              </div>
            </div>

            {/* 2행: 달력 + 선택일 타임라인 */}
            <DashboardCalendarPanel />

            {/* 3행: 미수금 + 할일 + 알림 */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
              <OutstandingCard outstanding={outstanding} />
              {cardVisibility.todos !== false && <TodoCard />}
              <AlertsCard
                lowStock={lowStock} waybill={waybillAlert} newReviews={newReviews}
                purchasing={purchasingAlert} supplies={suppliesAlert}
              />
            </div>
          </div>
        )}
      </div>
    </>
  );
}
