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
import { useQuery } from '@tanstack/react-query';
import Link from 'next/link';
import {
  ExternalLink,
  ShoppingCart, MessageSquare, Wrench, Store,
  AlertTriangle,
  PackageX, Truck, PackageOpen, Star, ClipboardList,
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
import { formatPhone, formatKRW, formatDate } from '@/lib/utils/format';
import { useActivityTypes } from '@/hooks/use-activity-types';
import { ActivityChips } from '@/components/shared/activity-chips';
import { NowTodoCard } from '@/components/dashboard/now-todo-card';

// 2026-11-15 금액 표기 통일 — 전엔 여기만 「₩142만」으로 줄여 보여 다른 화면(910,000원)과 달랐다
function fmtKRW(n: number) {
  return formatKRW(n || 0);
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
        className="flex-1 min-w-0 h-9 lg:h-8 px-3 rounded-lg border border-stone-200 text-sm focus:outline-none focus:border-stone-400 transition" /* 2026-10-19 모바일: min-w-0로 저장 버튼 밀림 방지 */
      />
      <button onClick={onSave} className="shrink-0 h-9 lg:h-auto px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition">저장</button>
      {isEdit && onCancel && (
        <button onClick={onCancel} className="shrink-0 h-9 lg:h-auto text-xs text-stone-400 hover:text-stone-600 transition">취소</button>
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
  // 2026-11-15 할 일은 작업실 한 곳에서 관리한다 → 여기서는 더 받지 않고 [작업실 열기]만 둔다.
  //   전에 적어 둔 메모는 지우지 않고 읽기만 할 수 있게 접어 둔다(작업실로 옮기기 전까지).
  const [open, setOpen] = useState(false);
  const { data } = useQuery({
    queryKey: ['dashboard-todos'],
    staleTime: 30_000,
    queryFn: async () => {
      const res = await fetch('/api/dashboard/todos');
      if (!res.ok) return { todos: [] };
      return res.json() as Promise<{ todos: Array<{ id: string; text: string; created_at: string }> }>;
    },
  });
  const todos = data?.todos || [];

  return (
    <div className="bg-white rounded-2xl border border-stone-200 p-4 h-full flex flex-col">
      <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold mb-2">할 일</p>
      <p className="text-xs text-stone-500 leading-relaxed">할 일은 「마모루 작업실」 한 곳에서 관리합니다.</p>
      <a
        href="https://claude.ai/artifact/NXww8NJQV6hKwayGHQHs5K" target="_blank" rel="noopener noreferrer"
        className="mt-3 inline-flex items-center justify-center gap-1.5 h-9 px-4 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition self-start"
      >
        작업실 열기 <ExternalLink size={12} />
      </a>
      {todos.length > 0 && (
        <div className="mt-3 pt-3 border-t border-stone-100">
          <button type="button" onClick={() => setOpen((o) => !o)} aria-expanded={open}
            className="text-[11px] font-semibold text-stone-500 hover:text-stone-900">
            전에 적어 둔 메모 {todos.length}개 {open ? '접기' : '보기'}
          </button>
          {open && (
            <ul className="mt-2 space-y-1.5">
              {todos.map((t) => (
                <li key={t.id} className="text-xs text-stone-700 leading-relaxed break-keep">
                  <span className="text-stone-400 tabular-nums mr-1.5">{formatDate(t.created_at)}</span>{t.text}
                </li>
              ))}
            </ul>
          )}
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
    // 2026-11-15 「운송장 1157」이 무엇의 수인지 알 수 없었다 → 롯데에서 받아 둔 송장 번호 중 남은 개수
    { icon: Truck,         label: '남은 송장 번호', count: waybillCount, href: '/settings',      active: waybillCount < 100, color: 'amber'  },
    { icon: Star,          label: '신규 후기', count: reviewCount,     href: '/reviews',        active: reviewCount > 0,    color: 'yellow' },
    { icon: PackageOpen,   label: '매입 대기', count: purchasingCount, href: '/purchasing',     active: purchasingCount > 0,color: 'blue'   },
    { icon: ClipboardList, label: '부자재',   count: suppliesCount,    href: '/supplies',       active: suppliesCount > 0,  color: 'stone'  },
  ];

  const shown = alerts.filter((a) => a.count > 0 || (a.label === '남은 송장 번호' && !!waybill));

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
      {/* 2026-11-15 0인 칸은 숨긴다 (남은 송장 번호는 0일 때가 가장 급하므로 항상 보인다) */}
      {shown.length === 0 && <p className="py-4 text-center text-xs text-stone-400">알릴 것이 없습니다</p>}
      <div className="grid gap-1.5" style={{ gridTemplateColumns: `repeat(${Math.max(shown.length, 2)}, minmax(0, 1fr))` }}>
        {shown.map((a) => {
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
  const { data: stats, isLoading, isError } = useHubStats();
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
        value={stats?.repairs.readyToShip || 0} primarySub="작업완료"
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
        {/* 장애 시 전부 "0"으로 보여 정상 오인 방지 — 명시적 경고 */}
        {isError && (
          <div className="mb-3 flex items-center gap-2 rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700">
            <AlertTriangle size={16} className="shrink-0" />
            데이터를 불러오지 못했습니다. 표시된 수치가 실제와 다를 수 있어요 — 잠시 후 새로고침해 주세요.
          </div>
        )}
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
            {/* 0행: 지금 처리할 것 (2026-11-15) */}
            <NowTodoCard />

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
