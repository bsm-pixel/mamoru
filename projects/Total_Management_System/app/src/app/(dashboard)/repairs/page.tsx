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
import { RefreshCw, Scissors, CreditCard, AlertTriangle, TrendingUp, Clapperboard } from 'lucide-react';
import { ContentCandidateListModal } from '@/components/repairs/content-candidate';
import { VisitBookingModal } from '@/components/repairs/visit-booking-modal';
import { formatKRW } from '@/lib/utils/format';
import { useListEscape } from '@/components/layout/list-keys';
import { useUrlSelection } from '@/hooks/use-url-selection';
import { RepairCustomerAside } from '@/components/repairs/repair-customer-aside';
import { emptyPaneCls } from '@/lib/ui/detail-pane';

export default function RepairsPage() {
  const sync = useRepairSync();
  const [selectedId, setSelectedId] = useUrlSelection(); // 2026-11-15 고른 건을 주소(?id=)에 — 뒤로 가기 = 상세 닫기
  useListEscape(() => { if (selectedId) setSelectedId(null); }); // 2026-11-15 Esc = 상세 닫기
  const [bookingOpen, setBookingOpen] = useState(false);   // 방문 예약 등록 모달
  const [contentOpen, setContentOpen] = useState(false);   // 콘텐츠 후보 목록
  const { data: stats } = useRepairDashboardStats();
  const [badgeFilter, setBadgeFilter] = useState<{ tab?: 'intake' | 'in_progress'; unpaidOnly?: boolean; staleOnly?: boolean } | null>(null);

  const isLg = useIsLg();

  // 상단 통계(매출 KPI + 상태 4카드) — 좌측 컬럼 상단에 배치(판매관리와 동일 IA: 상세가 우측 전체 높이 사용)
  // 2026-11-15 모바일: 매출 카드를 목록 아래로 내린다(전엔 카드와 칸이 먼저 나와 첫 접수 건이 화면 맨 아래에 걸렸다)
  //   max-lg:contents = 모바일에서는 이 묶음을 풀어 안의 두 덩어리가 목록과 같은 줄에서 순서를 바꿀 수 있게 한다
  const topStats = stats && (
    <div className="space-y-2.5 max-lg:contents lg:@container">
      {/* 매출 KPI + 오늘/이번주 (컴팩트) */}
      {(() => {
        const totalBags = (stats.monthRepairMamoru?.count || 0) + (stats.monthRepairOther?.count || 0) + (stats.monthRepairB2B?.count || 0);
        return (
          <div className="grid grid-cols-1 lg:grid-cols-[1fr_auto_auto] lg:@max-[620px]:grid-cols-2 lg:@max-[620px]:[&>*:first-child]:col-span-2 gap-2.5 max-lg:order-last">
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
            <div className="hidden lg:block bg-white rounded-2xl border border-stone-200 px-3 py-2.5 lg:w-36 lg:@max-[620px]:w-auto">
              <p className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">오늘 작업</p>
              <p className="text-xl font-bold text-stone-900 leading-tight">
                {(stats.todayWork?.mamoru || 0) + (stats.todayWork?.other || 0)}<span className="text-xs text-stone-500 ml-0.5">정</span>
              </p>
              <p className="text-[10px] text-stone-400">마모루 {stats.todayWork?.mamoru || 0} · 타사 {stats.todayWork?.other || 0}</p>
            </div>
            <div className="hidden lg:block bg-white rounded-2xl border border-stone-200 px-3 py-2.5 lg:w-36 lg:@max-[620px]:w-auto">
              <p className="text-[10px] uppercase tracking-wider text-stone-500 font-semibold">이번주 누적</p>
              <p className="text-xl font-bold text-stone-900 leading-tight">
                {(stats.weekWork?.mamoru || 0) + (stats.weekWork?.other || 0) + (stats.weekWork?.b2b || 0)}<span className="text-xs text-stone-500 ml-0.5">정</span>
              </p>
              <p className="text-[10px] text-stone-400">마모루 {stats.weekWork?.mamoru || 0} · 타사 {stats.weekWork?.other || 0}</p>
            </div>
          </div>
        );
      })()}

      {/* 상태 카드 — PC 4개. 모바일은 신규·진행을 뺀 2개
          (2026-10-19: 모바일은 목록 위 단계 버튼 판이 신규·진행 건수를 이미 보여줘 중복. 단계와 무관한 필터 2개만 남김) */}
      {/* 2026-11-15 PC: 신규 · 진행 카드는 아래 단계 탭과 같은 숫자라 뺐다. 단계와 무관한 「확인 필요 · 지연」만 탭 위에 작은 버튼으로 */}
      {isLg ? (
        ((stats.unpaidCount || 0) > 0 || (stats.staleCount || 0) > 0 || badgeFilter?.unpaidOnly || badgeFilter?.staleOnly) && (
          <div className="flex items-center gap-2">
            <span className="text-[11px] font-semibold text-stone-400">확인 필요</span>
            {((stats.unpaidCount || 0) > 0 || badgeFilter?.unpaidOnly) && (
              <button type="button" aria-pressed={!!badgeFilter?.unpaidOnly}
                onClick={() => setBadgeFilter(badgeFilter?.unpaidOnly ? null : { unpaidOnly: true })}
                className={`h-8 px-3 rounded-lg text-xs font-semibold border transition ${badgeFilter?.unpaidOnly ? 'bg-stone-900 text-white border-stone-900' : 'bg-white text-stone-700 border-stone-200 hover:bg-stone-50'}`}>
                미입금 {stats.unpaidCount || 0}
              </button>
            )}
            {((stats.staleCount || 0) > 0 || badgeFilter?.staleOnly) && (
              <button type="button" aria-pressed={!!badgeFilter?.staleOnly}
                onClick={() => setBadgeFilter(badgeFilter?.staleOnly ? null : { staleOnly: true })}
                className={`h-8 px-3 rounded-lg text-xs font-semibold border transition ${badgeFilter?.staleOnly ? 'bg-stone-900 text-white border-stone-900' : 'bg-white text-stone-700 border-stone-200 hover:bg-stone-50'}`}>
                3일 넘게 멈춤 {stats.staleCount || 0}
              </button>
            )}
          </div>
        )
      ) : (
      <div className="grid gap-2 grid-cols-2">
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
      )}
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
      {/* 2026-10-19 상단바는 새로고침만 — 모바일에선 아이콘만(글자가 2줄로 꺾이던 문제).
          [방문 예약 등록]은 작업 영역(검색줄 수거 불가일 옆)으로 이동 */}
      <Topbar title="복원수리" count={stats?.workingCount} hideRefresh action={
        <div className="flex items-center gap-2">
          {/* 2026-11-26 콘텐츠 후보 목록 (작업실 19번) */}
          <Button variant="secondary" size="sm" onClick={() => setContentOpen(true)} className="whitespace-nowrap" title="콘텐츠 후보 목록">
            <Clapperboard size={14} />
            <span className="hidden sm:inline">콘텐츠 후보</span>
          </Button>
          <Button variant="secondary" size="sm" onClick={() => sync.mutate()} disabled={sync.isPending} className="whitespace-nowrap" title="목록 새로고침">
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            <span className="hidden sm:inline">{sync.isPending ? '새로고침 중...' : '새로고침'}</span>
          </Button>
        </div>
      } />
      <VisitBookingModal open={bookingOpen} onClose={() => setBookingOpen(false)} onCreated={(id) => setSelectedId(id)} />
      <ContentCandidateListModal open={contentOpen} onClose={() => setContentOpen(false)} onSelect={(id) => setSelectedId(id)} />

      {isLg ? (
        /* PC: 좌측(통계+밀집그리드) + 우측 상세(전체 높이) — 판매관리와 동일 IA */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-[calc(100vh-92px)] bg-stone-50">
          {/* 2026-11-15 큰 화면: 표를 끝까지 늘리지 않고 칸을 더 쓴다
              ~1439 지금처럼 / 1440~ 목록(최대 1,100) + 상세 560 / 1920~ 상세 640 / 2200~ 목록(최대 900) + 상세 640 + 그 고객의 이력 480
              (왼쪽 메뉴 224 를 빼면 1920 에서는 3칸이 좁아 목록이 496 밖에 안 남는다 → 3칸은 2200 부터) */}
          <div className="flex-1 min-w-0 overflow-auto space-y-2.5 pr-1 min-[1440px]:max-w-[1100px] min-[2200px]:max-w-[900px]">
            {topStats}
            <RepairList
              onSelect={setSelectedId}
              selectedId={selectedId}
              initialTab={badgeFilter?.tab}
              unpaidOnly={badgeFilter?.unpaidOnly}
              staleOnly={badgeFilter?.staleOnly}
              onClearFilter={() => setBadgeFilter(null)}
              onBookVisit={() => setBookingOpen(true)}
              gridMode
            />
          </div>
          <div className={`w-[440px] min-[1440px]:w-[560px] min-[1920px]:w-[640px] shrink-0 overflow-y-auto ${emptyPaneCls(!!selectedId)}`}>
            {detailPanel}
          </div>
          <div className="hidden min-[2200px]:block w-[480px] shrink-0 overflow-y-auto">
            {selectedId ? (
              <div className="bg-white rounded-xl border border-neutral-200">
                <p className="px-4 pt-3 text-[11px] font-semibold text-neutral-400">이 고객의 이력 · 메모</p>
                <RepairCustomerAside repairId={selectedId} />
              </div>
            ) : null}
          </div>
        </div>
      ) : (
        /* 모바일: 통계 + 목록 + 슬라이드 패널 */
        <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 flex flex-col gap-3">
          {topStats}
          <RepairList
            onSelect={setSelectedId}
            selectedId={selectedId}
            initialTab={badgeFilter?.tab}
            unpaidOnly={badgeFilter?.unpaidOnly}
            staleOnly={badgeFilter?.staleOnly}
            onClearFilter={() => setBadgeFilter(null)}
            onBookVisit={() => setBookingOpen(true)}
          />
          <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="복원수리 상세">
            {selectedId && <RepairDetailPanel repairId={selectedId} />}
          </SlidePanel>
        </div>
      )}
    </>
  );
}
