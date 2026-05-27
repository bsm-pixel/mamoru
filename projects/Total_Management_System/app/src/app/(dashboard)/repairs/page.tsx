'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { RepairList } from '@/components/repairs/repair-list';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { StatCard } from '@/components/ui/stat-card';
import { RevenueDarkCard } from '@/components/ui/revenue-dark-card';
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

      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-4">
        {/* 1행: 이번달 매출(공통 RevenueDarkCard) + 오늘 작업 + 이번주 누적 */}
        {stats && (() => {
          const totalBags = (stats.monthRepairMamoru?.count || 0) + (stats.monthRepairOther?.count || 0) + (stats.monthRepairB2B?.count || 0);
          return (
            <div className="grid grid-cols-1 lg:grid-cols-[1fr_auto_auto] gap-3">
              <RevenueDarkCard
                icon={TrendingUp}
                label="이번달 복원수리"
                amount={formatKRW(stats.monthRepairAmount)}
                rightValue={String(totalBags)}
                splits={[
                  { label: '마모루', amount: formatKRW(stats.monthRepairMamoru?.amount || 0), sub: `${stats.monthRepairMamoru?.count || 0}자루` },
                  { label: '타사',   amount: formatKRW(stats.monthRepairOther?.amount  || 0), sub: `${stats.monthRepairOther?.count  || 0}자루` },
                  { label: 'B2B',    amount: formatKRW(stats.monthRepairB2B?.amount    || 0), sub: `${stats.monthRepairB2B?.count    || 0}자루` },
                ]}
              />

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
          );
        })()}

        {/* 2행: 상태 4카드 (공통 StatCard) + 새로고침 */}
        <div className="grid grid-cols-12 gap-3">
          <div className="col-span-12 lg:col-span-9 grid grid-cols-2 sm:grid-cols-4 gap-3">
            <StatCard label="신규" icon={Inbox} accent="blue"
              value={stats?.intakeNew || 0} primarySub="확인 필요"
              active={badgeFilter?.tab === 'intake'}
              onClick={() => setBadgeFilter(badgeFilter?.tab === 'intake' ? null : { tab: 'intake' })}
            />
            <StatCard label="진행" icon={Loader} accent="amber"
              value={stats?.workingCount || 0} primarySub="작업 중"
              active={badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly}
              onClick={() => setBadgeFilter(badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly ? null : { tab: 'in_progress' })}
            />
            <StatCard label="미입금" icon={CreditCard} accent="rose" dimWhenZero
              value={stats?.unpaidCount || 0} primarySub="확인 필요"
              active={badgeFilter?.unpaidOnly}
              onClick={() => setBadgeFilter(badgeFilter?.unpaidOnly ? null : { unpaidOnly: true })}
            />
            <StatCard label="3일경과" icon={AlertTriangle} accent="orange" dimWhenZero
              value={stats?.staleCount || 0} primarySub="지연"
              active={badgeFilter?.staleOnly}
              onClick={() => setBadgeFilter(badgeFilter?.staleOnly ? null : { staleOnly: true })}
            />
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
