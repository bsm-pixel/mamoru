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
import { useRepairTabData } from '@/hooks/use-repair-tabs';
import { AlertTriangle } from 'lucide-react';

export default function RepairDashboardPage() {
  const [activeTab, setActiveTab] = useState<RepairTabKey>('intake');
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
        <IntakeTab repairs={tabData.intake} isLoading={isLoading} />
      )}
      {activeTab === 'pickup_needed' && (
        <PickupNeededTab repairs={tabData.pickup_needed} isLoading={isLoading} />
      )}
      {activeTab === 'inbound_waiting' && (
        <InboundWaitingTab repairs={tabData.inbound_waiting} isLoading={isLoading} />
      )}
      {activeTab === 'in_progress' && (
        <InProgressTab repairs={tabData.in_progress} isLoading={isLoading} />
      )}
      {activeTab === 'ready_to_ship' && (
        <ReadyToShipTab repairs={tabData.ready_to_ship} isLoading={isLoading} />
      )}
      {activeTab === 'shipped' && (
        <ShippedTab repairs={tabData.shipped} isLoading={isLoading} />
      )}
    </>
  );
}
