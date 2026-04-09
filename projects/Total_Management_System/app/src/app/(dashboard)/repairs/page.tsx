'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { RepairList } from '@/components/repairs/repair-list';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useRepairSync } from '@/hooks/use-repairs';
import { useRepairDashboardStats } from '@/hooks/use-dashboard-stats';
import { RefreshCw, Scissors, Inbox, Loader, CreditCard, AlertTriangle, TrendingUp } from 'lucide-react';
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

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 이번달 매출 + 오늘 작업 + 이번주 누적 — PC 3컬럼 / 모바일 세로 */}
        {stats && (
          <div className="grid grid-cols-1 lg:grid-cols-[1fr_auto_auto] gap-3">
            {/* 이번달 복원수리 매출 */}
            <div className="rounded-lg bg-neutral-900 text-white overflow-hidden">
              <div className="flex items-center gap-3 px-4 py-3">
                <TrendingUp size={18} className="opacity-70" />
                <div>
                  <p className="text-xs opacity-70">이번달 복원수리 매출</p>
                  <p className="text-lg font-bold">{formatKRW(stats.monthRepairAmount)}</p>
                </div>
                <p className="ml-auto text-lg font-bold">
                  {(stats.monthRepairMamoru?.count || 0) + (stats.monthRepairOther?.count || 0) + (stats.monthRepairB2B?.count || 0)}건
                </p>
              </div>
              <div className="grid grid-cols-3 gap-px bg-white/10">
                <div className="px-3 py-2 text-center">
                  <p className="text-[10px] opacity-50">마모루</p>
                  <p className="text-sm font-bold">{formatKRW(stats.monthRepairMamoru?.amount || 0)}</p>
                  <p className="text-[10px] opacity-40">{stats.monthRepairMamoru?.count || 0}건</p>
                </div>
                <div className="px-3 py-2 text-center">
                  <p className="text-[10px] opacity-50">타사</p>
                  <p className="text-sm font-bold">{formatKRW(stats.monthRepairOther?.amount || 0)}</p>
                  <p className="text-[10px] opacity-40">{stats.monthRepairOther?.count || 0}건</p>
                </div>
                <div className="px-3 py-2 text-center">
                  <p className="text-[10px] opacity-50">B2B</p>
                  <p className="text-sm font-bold">{formatKRW(stats.monthRepairB2B?.amount || 0)}</p>
                  <p className="text-[10px] opacity-40">{stats.monthRepairB2B?.count || 0}건</p>
                </div>
              </div>
            </div>

            {/* 오늘 작업 완료 */}
            <div className="bg-white rounded-lg border border-neutral-200 p-3 lg:w-44">
              <p className="text-xs text-neutral-500 mb-1">오늘 작업 완료</p>
              <p className="text-lg font-bold text-neutral-900">
                {(stats.todayWork?.mamoru || 0) + (stats.todayWork?.other || 0)}정
              </p>
              <p className="text-[11px] text-neutral-400">
                마모루 {stats.todayWork?.mamoru || 0} · 타사 {stats.todayWork?.other || 0} ({stats.todayWork?.count || 0}건)
              </p>
            </div>

            {/* 이번주 누적 */}
            <div className="bg-white rounded-lg border border-neutral-200 p-3 lg:w-44">
              <p className="text-xs text-neutral-500 mb-1">이번주 누적</p>
              <p className="text-lg font-bold text-neutral-900">
                {(stats.weekWork?.mamoru || 0) + (stats.weekWork?.other || 0)}정
              </p>
              <p className="text-[11px] text-neutral-400">
                마모루 {stats.weekWork?.mamoru || 0} · 타사 {stats.weekWork?.other || 0} ({stats.weekWork?.count || 0}건)
              </p>
            </div>
          </div>
        )}

        {/* 상단: 요약 카드 + 새로고침 */}
        <div className="flex items-center gap-3">
          <div className="flex gap-2 flex-1 min-w-0">
            <button onClick={() => setBadgeFilter(badgeFilter?.tab === 'intake' ? null : { tab: 'intake' })}
              className={`flex items-center gap-2 px-3 py-2 rounded-lg transition min-w-0 ${badgeFilter?.tab === 'intake' ? 'bg-blue-200 ring-2 ring-blue-400' : 'bg-blue-50 hover:bg-blue-100'}`}>
              <Inbox size={14} className="text-blue-600 shrink-0" />
              <span className="text-xs text-neutral-500">신규</span>
              <span className="text-sm font-bold text-blue-700">{stats?.intakeNew || 0}</span>
            </button>
            <button onClick={() => setBadgeFilter(badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly ? null : { tab: 'in_progress' })}
              className={`flex items-center gap-2 px-3 py-2 rounded-lg transition min-w-0 ${badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly ? 'bg-amber-200 ring-2 ring-amber-400' : 'bg-amber-50 hover:bg-amber-100'}`}>
              <Loader size={14} className="text-amber-600 shrink-0" />
              <span className="text-xs text-neutral-500">진행</span>
              <span className="text-sm font-bold text-amber-700">{stats?.workingCount || 0}</span>
            </button>
            {(stats?.unpaidCount || 0) > 0 && (
              <button onClick={() => setBadgeFilter(badgeFilter?.unpaidOnly ? null : { unpaidOnly: true })}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg transition min-w-0 ${badgeFilter?.unpaidOnly ? 'bg-red-200 ring-2 ring-red-400' : 'bg-red-50 hover:bg-red-100'}`}>
                <CreditCard size={14} className="text-red-600 shrink-0" />
                <span className="text-xs text-neutral-500">미입금</span>
                <span className="text-sm font-bold text-red-700">{stats?.unpaidCount || 0}</span>
              </button>
            )}
            {(stats?.staleCount || 0) > 0 && (
              <button onClick={() => setBadgeFilter(badgeFilter?.staleOnly ? null : { staleOnly: true })}
                className={`flex items-center gap-2 px-3 py-2 rounded-lg transition min-w-0 ${badgeFilter?.staleOnly ? 'bg-orange-200 ring-2 ring-orange-400' : 'bg-orange-50 hover:bg-orange-100'}`}>
                <AlertTriangle size={14} className="text-orange-600 shrink-0" />
                <span className="text-xs text-neutral-500">3일경과</span>
                <span className="text-sm font-bold text-orange-700">{stats?.staleCount || 0}</span>
              </button>
            )}
          </div>
          <Button
            variant="secondary"
            size="sm"
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
            className="shrink-0"
          >
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '새로고침 중...' : '새로고침'}
          </Button>
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
