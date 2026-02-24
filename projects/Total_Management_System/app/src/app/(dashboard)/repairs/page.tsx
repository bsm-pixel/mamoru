'use client';

import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { RepairList } from '@/components/repairs/repair-list';
import { useRepairSync } from '@/hooks/use-repairs';
import { RefreshCw } from 'lucide-react';

export default function RepairsPage() {
  const sync = useRepairSync();

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

        {/* 목록 */}
        <RepairList />
      </div>
    </>
  );
}
