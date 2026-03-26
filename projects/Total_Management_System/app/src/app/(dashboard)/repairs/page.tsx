'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { RepairList } from '@/components/repairs/repair-list';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useRepairSync } from '@/hooks/use-repairs';
import { useRepairDashboardStats } from '@/hooks/use-dashboard-stats';
import { RefreshCw, Scissors, Inbox, Loader, CreditCard, AlertTriangle } from 'lucide-react';

export default function RepairsPage() {
  const sync = useRepairSync();
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const { data: stats } = useRepairDashboardStats();

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
        {/* 상단: 요약 카드 + 새로고침 */}
        <div className="flex items-center gap-3">
          <div className="flex gap-2 flex-1 min-w-0">
            <button className="flex items-center gap-2 px-3 py-2 rounded-lg bg-blue-50 hover:bg-blue-100 transition min-w-0">
              <Inbox size={14} className="text-blue-600 shrink-0" />
              <span className="text-xs text-neutral-500">신규</span>
              <span className="text-sm font-bold text-blue-700">{stats?.intakeNew || 0}</span>
            </button>
            <button className="flex items-center gap-2 px-3 py-2 rounded-lg bg-amber-50 hover:bg-amber-100 transition min-w-0">
              <Loader size={14} className="text-amber-600 shrink-0" />
              <span className="text-xs text-neutral-500">진행</span>
              <span className="text-sm font-bold text-amber-700">{stats?.workingCount || 0}</span>
            </button>
            {(stats?.unpaidCount || 0) > 0 && (
              <button className="flex items-center gap-2 px-3 py-2 rounded-lg bg-red-50 hover:bg-red-100 transition min-w-0">
                <CreditCard size={14} className="text-red-600 shrink-0" />
                <span className="text-xs text-neutral-500">미입금</span>
                <span className="text-sm font-bold text-red-700">{stats?.unpaidCount || 0}</span>
              </button>
            )}
            {(stats?.staleCount || 0) > 0 && (
              <button className="flex items-center gap-2 px-3 py-2 rounded-lg bg-orange-50 hover:bg-orange-100 transition min-w-0">
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
              <RepairList onSelect={setSelectedId} selectedId={selectedId} />
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
          <RepairList onSelect={setSelectedId} selectedId={selectedId} />
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
