'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { RepairTabBar, type RepairTabKey } from '@/components/repairs/repair-tab-bar';
import { IntakeTab } from '@/components/repairs/tabs/intake-tab';
import { PickupNeededTab } from '@/components/repairs/tabs/pickup-needed-tab';
import { InboundWaitingTab } from '@/components/repairs/tabs/inbound-waiting-tab';
import { InProgressTab } from '@/components/repairs/tabs/in-progress-tab';
import { ReadyToShipTab } from '@/components/repairs/tabs/ready-to-ship-tab';
import { ShippedTab } from '@/components/repairs/tabs/shipped-tab';
import { RepairDetailPanel } from '@/components/repairs/repair-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useRepairTabData } from '@/hooks/use-repair-tabs';
import { AlertTriangle } from 'lucide-react';

export default function RepairDashboardPage() {
  const [activeTab, setActiveTab] = useState<RepairTabKey>('intake');
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const { tabs, tabData, isLoading, staleCount } = useRepairTabData();

  return (
    <>
      <Topbar title="복원수리 대시보드" />

      {/* 경과일 3일 이상 경고 배너 */}
      {staleCount > 0 && (
        <div className="mx-4 md:mx-6 mt-3 flex items-center gap-2 bg-warning/10 border border-warning/30 rounded-lg px-4 py-3">
          <AlertTriangle size={18} className="text-warning shrink-0" />
          <p className="text-sm text-warning font-medium">
            접수/비용안내 후 <span className="font-bold">{staleCount}건</span>이 3일 이상 미처리
          </p>
        </div>
      )}

      {/* 고정 탭 바 */}
      <RepairTabBar
        tabs={tabs}
        activeTab={activeTab}
        onTabChange={setActiveTab}
      />

      {/* 탭별 콘텐츠 */}
      {activeTab === 'intake' && (
        <IntakeTab repairs={tabData.intake} isLoading={isLoading} onSelect={setSelectedId} />
      )}
      {activeTab === 'pickup_needed' && (
        <PickupNeededTab repairs={tabData.pickup_needed} isLoading={isLoading} onSelect={setSelectedId} />
      )}
      {activeTab === 'inbound_waiting' && (
        <InboundWaitingTab repairs={tabData.inbound_waiting} isLoading={isLoading} onSelect={setSelectedId} />
      )}
      {activeTab === 'in_progress' && (
        <InProgressTab repairs={tabData.in_progress} isLoading={isLoading} onSelect={setSelectedId} />
      )}
      {activeTab === 'ready_to_ship' && (
        <ReadyToShipTab repairs={tabData.ready_to_ship} isLoading={isLoading} onSelect={setSelectedId} />
      )}
      {activeTab === 'shipped' && (
        <ShippedTab repairs={tabData.shipped} isLoading={isLoading} onSelect={setSelectedId} />
      )}

      {/* 상세 슬라이드 패널 */}
      <SlidePanel
        open={!!selectedId}
        onClose={() => setSelectedId(null)}
        title="복원수리 상세"
        className="sm:w-[640px]"
      >
        {selectedId && <RepairDetailPanel repairId={selectedId} />}
      </SlidePanel>
    </>
  );
}
