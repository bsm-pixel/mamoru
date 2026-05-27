'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { RepairList } from '@/components/repairs/repair-list';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useRepairSync } from '@/hooks/use-repairs';
import { useRepairDashboardStats } from '@/hooks/use-dashboard-stats';
import { RefreshCw, Scissors, Inbox, Loader, CreditCard, AlertTriangle, TrendingUp, ArrowRight } from 'lucide-react';
import { formatKRW } from '@/lib/utils/format';

export default function RepairsPage() {
  const sync = useRepairSync();
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const { data: stats } = useRepairDashboardStats();
  const [badgeFilter, setBadgeFilter] = useState<{ tab?: 'intake' | 'in_progress'; unpaidOnly?: boolean; staleOnly?: boolean } | null>(null);

  // PC 여부 감지 (lg:1024px+) — isLg conditional rendering
  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  return (
    <>
      <Topbar title="복원수리" />

      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-4">
        {/* 1행: 이번달 매출(A2 그라데이션) + 오늘 작업 + 이번주 누적 */}
        {stats && (
          <div className="grid grid-cols-1 lg:grid-cols-[1fr_auto_auto] gap-3">
            {/* 매출 KPI (A2 톤 — 사장님 채택: 그라데이션 부드러운 다크) */}
            <div className="rounded-2xl bg-gradient-to-br from-stone-800 to-stone-900 text-white overflow-hidden ring-1 ring-white/5">
              <div className="flex items-center gap-3 px-5 py-4">
                <TrendingUp size={18} className="opacity-60" />
                <div>
                  <p className="text-[11px] uppercase tracking-wider opacity-60 font-semibold">이번달 복원수리</p>
                  <p className="text-2xl font-bold">{formatKRW(stats.monthRepairAmount)}</p>
                </div>
                <p className="ml-auto text-xl font-bold">
                  {(stats.monthRepairMamoru?.count || 0) + (stats.monthRepairOther?.count || 0) + (stats.monthRepairB2B?.count || 0)}<span className="text-xs opacity-60 ml-0.5">자루</span>
                </p>
              </div>
              <div className="grid grid-cols-3 gap-px bg-white/10">
                <div className="px-3 py-2.5 text-center bg-stone-900/40">
                  <p className="text-[10px] opacity-50 uppercase tracking-wider">마모루</p>
                  <p className="text-sm font-bold mt-0.5">{formatKRW(stats.monthRepairMamoru?.amount || 0)}</p>
                  <p className="text-[10px] opacity-40">{stats.monthRepairMamoru?.count || 0}자루</p>
                </div>
                <div className="px-3 py-2.5 text-center bg-stone-900/40">
                  <p className="text-[10px] opacity-50 uppercase tracking-wider">타사</p>
                  <p className="text-sm font-bold mt-0.5">{formatKRW(stats.monthRepairOther?.amount || 0)}</p>
                  <p className="text-[10px] opacity-40">{stats.monthRepairOther?.count || 0}자루</p>
                </div>
                <div className="px-3 py-2.5 text-center bg-stone-900/40">
                  <p className="text-[10px] opacity-50 uppercase tracking-wider">B2B</p>
                  <p className="text-sm font-bold mt-0.5">{formatKRW(stats.monthRepairB2B?.amount || 0)}</p>
                  <p className="text-[10px] opacity-40">{stats.monthRepairB2B?.count || 0}자루</p>
                </div>
              </div>
            </div>

            {/* 오늘 작업 완료 — PC만 표시 */}
            <div className="hidden lg:block bg-white rounded-2xl border border-stone-200 p-4 lg:w-44">
              <p className="text-[11px] uppercase tracking-wider text-stone-500 font-semibold mb-1">오늘 작업</p>
              <p className="text-2xl font-bold text-stone-900">
                {(stats.todayWork?.mamoru || 0) + (stats.todayWork?.other || 0)}<span className="text-xs text-stone-500 ml-0.5">정</span>
              </p>
              <p className="text-[10px] text-stone-400 mt-0.5">
                마모루 {stats.todayWork?.mamoru || 0} · 타사 {stats.todayWork?.other || 0} ({stats.todayWork?.count || 0}건)
              </p>
            </div>

            {/* 이번주 누적 — PC만 표시 */}
            <div className="hidden lg:block bg-white rounded-2xl border border-stone-200 p-4 lg:w-44">
              <p className="text-[11px] uppercase tracking-wider text-stone-500 font-semibold mb-1">이번주 누적</p>
              <p className="text-2xl font-bold text-stone-900">
                {(stats.weekWork?.mamoru || 0) + (stats.weekWork?.other || 0) + (stats.weekWork?.b2b || 0)}<span className="text-xs text-stone-500 ml-0.5">정</span>
              </p>
              <p className="text-[10px] text-stone-400 mt-0.5">
                마모루 {stats.weekWork?.mamoru || 0} · 타사 {stats.weekWork?.other || 0}{(stats.weekWork?.b2b || 0) > 0 ? ` · B2B ${stats.weekWork.b2b}` : ''} ({stats.weekWork?.count || 0}건)
              </p>
            </div>
          </div>
        )}

        {/* 2행: 상태 4카드 (시안 B CategoryCard 톤) + 새로고침 */}
        <div className="grid grid-cols-12 gap-3">
          <div className="col-span-12 lg:col-span-9 grid grid-cols-2 sm:grid-cols-4 gap-3">
            {/* 신규 */}
            <button
              onClick={() => setBadgeFilter(badgeFilter?.tab === 'intake' ? null : { tab: 'intake' })}
              className={`bg-white rounded-2xl border p-4 hover:border-stone-300 transition group text-left ${badgeFilter?.tab === 'intake' ? 'border-stone-900 ring-1 ring-stone-900' : 'border-stone-200'}`}
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <Inbox size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">신규</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className="text-3xl font-bold leading-none text-blue-600">{stats?.intakeNew || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">확인 필요</p>
            </button>
            {/* 진행 */}
            <button
              onClick={() => setBadgeFilter(badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly ? null : { tab: 'in_progress' })}
              className={`bg-white rounded-2xl border p-4 hover:border-stone-300 transition group text-left ${badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly ? 'border-stone-900 ring-1 ring-stone-900' : 'border-stone-200'}`}
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <Loader size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">진행</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className="text-3xl font-bold leading-none text-amber-600">{stats?.workingCount || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">작업 중</p>
            </button>
            {/* 미입금 */}
            <button
              onClick={() => setBadgeFilter(badgeFilter?.unpaidOnly ? null : { unpaidOnly: true })}
              className={`bg-white rounded-2xl border p-4 hover:border-stone-300 transition group text-left ${badgeFilter?.unpaidOnly ? 'border-stone-900 ring-1 ring-stone-900' : 'border-stone-200'}`}
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <CreditCard size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">미입금</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className={`text-3xl font-bold leading-none ${(stats?.unpaidCount || 0) > 0 ? 'text-rose-600' : 'text-stone-300'}`}>{stats?.unpaidCount || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">확인 필요</p>
            </button>
            {/* 3일경과 */}
            <button
              onClick={() => setBadgeFilter(badgeFilter?.staleOnly ? null : { staleOnly: true })}
              className={`bg-white rounded-2xl border p-4 hover:border-stone-300 transition group text-left ${badgeFilter?.staleOnly ? 'border-stone-900 ring-1 ring-stone-900' : 'border-stone-200'}`}
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <AlertTriangle size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">3일경과</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className={`text-3xl font-bold leading-none ${(stats?.staleCount || 0) > 0 ? 'text-orange-600' : 'text-stone-300'}`}>{stats?.staleCount || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">지연</p>
            </button>
          </div>
          <div className="col-span-12 lg:col-span-3 flex items-stretch">
            <button
              onClick={() => sync.mutate()}
              disabled={sync.isPending}
              className="w-full rounded-2xl border border-stone-200 bg-white hover:bg-stone-50 transition flex items-center justify-center gap-2 text-xs font-semibold text-stone-700 disabled:opacity-60 px-3 py-3"
            >
              <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
              {sync.isPending ? '새로고침 중...' : '새로고침'}
            </button>
          </div>
        </div>

        {/* PC: 좌측 목록 + 우측 상세 모니터 */}
        {isLg && (
          <div className="flex gap-4 h-[calc(100vh-220px)]">
            {/* 좌측: 목록 (2/5) */}
            <div className="w-[40%] shrink-0 overflow-y-auto">
              <RepairList
                onSelect={setSelectedId}
                selectedId={selectedId}
                initialTab={badgeFilter?.tab}
                unpaidOnly={badgeFilter?.unpaidOnly}
                staleOnly={badgeFilter?.staleOnly}
                onClearFilter={() => setBadgeFilter(null)}
              />
            </div>

            {/* 우측: 상세 모니터 (3/5) */}
            <div className="flex-1 min-w-0 overflow-y-auto">
              {selectedId ? (
                <RepairDetailPanel repairId={selectedId} />
              ) : (
                <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                  <Scissors size={28} className="mb-2 opacity-40" />
                  <p className="text-xs text-center">목록에서 복원수리 건을<br />선택하세요</p>
                </div>
              )}
            </div>
          </div>
        )}

        {/* 모바일: 목록만 */}
        {!isLg && (
          <RepairList
            onSelect={setSelectedId}
            selectedId={selectedId}
            initialTab={badgeFilter?.tab}
            unpaidOnly={badgeFilter?.unpaidOnly}
            staleOnly={badgeFilter?.staleOnly}
            onClearFilter={() => setBadgeFilter(null)}
          />
        )}

        {/* 모바일 전용 슬라이드 패널 */}
        {!isLg && (
          <SlidePanel
            open={!!selectedId}
            onClose={() => setSelectedId(null)}
            title="복원수리 상세"
          >
            {selectedId && <RepairDetailPanel repairId={selectedId} />}
          </SlidePanel>
        )}
      </div>
    </>
  );
}
