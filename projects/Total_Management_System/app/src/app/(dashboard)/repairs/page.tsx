'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { RepairList } from '@/components/repairs/repair-list';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useRepairSync } from '@/hooks/use-repairs';
import { RefreshCw, Scissors } from 'lucide-react';

export default function RepairsPage() {
  const sync = useRepairSync();
  const [selectedId, setSelectedId] = useState<string | null>(null);

  return (
    <>
      <Topbar title="복원수리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 상단: 새로고침 */}
        <div className="flex items-center justify-between">
          <Button
            variant="secondary"
            size="sm"
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
          >
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '새로고침 중...' : '새로고침'}
          </Button>
        </div>

        {/* PC: 좌측 목록 + 우측 상세 패널 / 모바일: 목록만 */}
        <div className="flex gap-6">
          {/* 좌측: 목록 */}
          <div className="flex-1 min-w-0 lg:max-w-[480px]">
            <RepairList
              onSelect={setSelectedId}
              selectedId={selectedId}
            />
          </div>

          {/* 우측: 상세 패널 (PC lg+ 만 표시) */}
          <div className="hidden lg:flex lg:flex-col flex-1 min-w-0 h-[calc(100vh-140px)]">
            {selectedId ? (
              <RepairDetailPanel repairId={selectedId} />
            ) : (
              <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                <Scissors size={32} className="mb-2 opacity-50" />
                <p className="text-sm">목록에서 복원수리 건을 선택하세요</p>
              </div>
            )}
          </div>
        </div>

        {/* 모바일: 슬라이드 패널로 상세 표시 */}
        <SlidePanel
          open={!!selectedId}
          onClose={() => setSelectedId(null)}
          title="복원수리 상세"
          className="lg:hidden w-full"
        >
          {selectedId && <RepairDetailPanel repairId={selectedId} />}
        </SlidePanel>
      </div>
    </>
  );
}
