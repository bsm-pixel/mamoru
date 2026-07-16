'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { RepairList } from '@/components/repairs/repair-list';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { StatCard } from '@/components/ui/stat-card';
import { RevenueDarkCard } from '@/components/ui/revenue-dark-card';
import { Button } from '@/components/ui/button';
import { useIsLg } from '@/hooks/use-grid-mode';
import { useRepairSync } from '@/hooks/use-repairs';
import { useRepairDashboardStats } from '@/hooks/use-dashboard-stats';
import { RefreshCw, Scissors, Inbox, Loader, CreditCard, AlertTriangle, TrendingUp } from 'lucide-react';
import { formatKRW } from '@/lib/utils/format';

export default function RepairsPage() {
  const sync = useRepairSync();
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const { data: stats } = useRepairDashboardStats();
  const [badgeFilter, setBadgeFilter] = useState<{ tab?: 'intake' | 'in_progress'; unpaidOnly?: boolean; staleOnly?: boolean } | null>(null);

  const isLg = useIsLg();

  // 상단 통계(매출 KPI + 상태 4카드) — 좌측 컬럼 상단에 배치(판매관리와 동일 IA: 상세가 우측 전체 높이 사용)
  const topStats = stats && (
    <div className="space-y-2.5">
      {/* 매출 KPI + 오늘/이번주 (컴팩트) */}
      {(() => {
        const totalBags = (stats.monthRepairMamoru?.count || 0) + (stats.monthRepairOther?.count || 0) + (stats.monthRepairB2B?.count || 0);
        return (
          <div className="grid grid-cols-1 lg:grid-cols-[1fr_auto_auto] gap-2.5">
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
            {/* 오늘/이번주 작업 — PC만, 컴팩트 */}
            <div className="hidden lg:block bg-white rounded-2xl border border-stone-200 px-3 py-2.5 lg:w-36">
              <p className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">오늘 작업</p>
              <p className="text-xl font-bold text-stone-900 leading-tight">
                {(stats.todayWork?.mamoru || 0) + (stats.todayWork?.other || 0)}<span className="text-xs text-stone-500 ml-0.5">정</span>
              </p>
              <p className="text-[10px] text-stone-400">마모루 {stats.todayWork?.mamoru || 0} · 타사 {stats.todayWork?.other || 0}</p>
            </div>
            <div className="hidden lg:block bg-white rounded-2xl border border-stone-200 px-3 py-2.5 lg:w-36">
              <p className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">이번주 누적</p>
              <p className="text-xl font-bold text-stone-900 leading-tight">
                {(stats.weekWork?.mamoru || 0) + (stats.weekWork?.other || 0) + (stats.weekWork?.b2b || 0)}<span className="text-xs text-stone-500 ml-0.5">정</span>
              </p>
              <p className="text-[10px] text-stone-400">마모루 {stats.weekWork?.mamoru || 0} · 타사 {stats.weekWork?.other || 0}</p>
            </div>
          </div>
        );
      })()}

      {/* 상태 4카드 (컴팩트 — 밀집) */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <StatCard compact label="신규" icon={Inbox} accent="blue"
          value={stats.intakeNew || 0} primarySub="확인 필요"
          active={badgeFilter?.tab === 'intake'}
          onClick={() => setBadgeFilter(badgeFilter?.tab === 'intake' ? null : { tab: 'intake' })}
        />
        <StatCard compact label="진행" icon={Loader} accent="amber"
          value={stats.workingCount || 0} primarySub="작업 중"
          active={badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly}
          onClick={() => setBadgeFilter(badgeFilter?.tab === 'in_progress' && !badgeFilter.unpaidOnly ? null : { tab: 'in_progress' })}
        />
        <StatCard compact label="미입금" icon={CreditCard} accent="rose" dimWhenZero
          value={stats.unpaidCount || 0} primarySub="확인 필요"
          active={badgeFilter?.unpaidOnly}
          onClick={() => setBadgeFilter(badgeFilter?.unpaidOnly ? null : { unpaidOnly: true })}
        />
        <StatCard compact label="3일경과" icon={AlertTriangle} accent="orange" dimWhenZero
          value={stats.staleCount || 0} primarySub="지연"
          active={badgeFilter?.staleOnly}
          onClick={() => setBadgeFilter(badgeFilter?.staleOnly ? null : { staleOnly: true })}
        />
      </div>
    </div>
  );

  const detailPanel = selectedId ? (
    <RepairDetailPanel repairId={selectedId} />
  ) : (
    <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
      <Scissors size={28} className="mb-2 opacity-40" />
      <p className="text-xs text-center">목록에서 복원수리 건을<br />선택하세요</p>
    </div>
  );

  return (
    <>
      <Topbar title="복원수리" action={
        <Button variant="secondary" size="sm" onClick={() => sync.mutate()} disabled={sync.isPending}>
          <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
          {sync.isPending ? '새로고침 중...' : '새로고침'}
        </Button>
      } />

      {isLg ? (
        /* PC: 좌측(통계+밀집그리드) + 우측 상세(전체 높이) — 판매관리와 동일 IA */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-[calc(100vh-92px)] bg-stone-50">
          <div className="flex-1 min-w-0 overflow-auto space-y-2.5 pr-1">
            {topStats}
            <RepairList
              onSelect={setSelectedId}
              selectedId={selectedId}
              initialTab={badgeFilter?.tab}
              unpaidOnly={badgeFilter?.unpaidOnly}
              staleOnly={badgeFilter?.staleOnly}
              onClearFilter={() => setBadgeFilter(null)}
              gridMode
            />
          </div>
          <div className="w-[440px] shrink-0 overflow-y-auto">
            {detailPanel}
          </div>
        </div>
      ) : (
        /* 모바일: 통계 + 목록 + 슬라이드 패널 */
        <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-3">
          {topStats}
          <RepairList
            onSelect={setSelectedId}
            selectedId={selectedId}
            initialTab={badgeFilter?.tab}
            unpaidOnly={badgeFilter?.unpaidOnly}
            staleOnly={badgeFilter?.staleOnly}
            onClearFilter={() => setBadgeFilter(null)}
          />
          <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="복원수리 상세">
            {selectedId && <RepairDetailPanel repairId={selectedId} />}
          </SlidePanel>
        </div>
      )}
    </>
  );
}
