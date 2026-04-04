'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { useSettings, useUpdateSettings } from '@/hooks/use-settings';
import {
  LayoutDashboard, Truck, MessageSquare, Wrench, Store,
  Users, Package, BarChart3, Bell, Settings, ChevronDown,
} from 'lucide-react';

/* ── 탭별 컴포넌트 ─────────────────────────── */
import DashboardSettings from '@/components/settings/dashboard-settings';
import ShippingSettings from '@/components/settings/shipping-settings';
import ConsultationSettings from '@/components/settings/consultation-settings';
import RepairSettings from '@/components/settings/repair-settings';
import SalesSettings from '@/components/settings/sales-settings';
import CustomerSettings from '@/components/settings/customer-settings';
import InventorySettings from '@/components/settings/inventory-settings';
import AccountingSettings from '@/components/settings/accounting-settings';
import NotificationSettings from '@/components/settings/notification-settings';
import SystemSettings from '@/components/settings/system-settings';

const TABS = [
  { key: 'dashboard', label: '대시보드', icon: LayoutDashboard },
  { key: 'shipping', label: '주문·배송', icon: Truck },
  { key: 'consultation', label: '상담', icon: MessageSquare },
  { key: 'repair', label: '복원수리', icon: Wrench },
  { key: 'sales', label: '판매', icon: Store },
  { key: 'customer', label: '고객', icon: Users },
  { key: 'inventory', label: '상품·재고', icon: Package },
  { key: 'accounting', label: '회계', icon: BarChart3 },
  { key: 'notifications', label: '알림·연동', icon: Bell },
  { key: 'system', label: '시스템', icon: Settings },
] as const;

type TabKey = (typeof TABS)[number]['key'];

const TAB_COMPONENTS: Record<TabKey, React.ComponentType<TabProps>> = {
  dashboard: DashboardSettings,
  shipping: ShippingSettings,
  consultation: ConsultationSettings,
  repair: RepairSettings,
  sales: SalesSettings,
  customer: CustomerSettings,
  inventory: InventorySettings,
  accounting: AccountingSettings,
  notifications: NotificationSettings,
  system: SystemSettings,
};

export interface TabProps {
  settings: Record<string, unknown>;
  onSave: (items: { key: string; value: unknown }[]) => void;
  saving: boolean;
}

export default function SettingsPage() {
  const [activeTab, setActiveTab] = useState<TabKey>('dashboard');
  const [isMd, setIsMd] = useState(false);
  const { data: settings, isLoading } = useSettings();
  const updateSettings = useUpdateSettings();

  useEffect(() => {
    const mq = window.matchMedia('(min-width: 768px)');
    setIsMd(mq.matches);
    const handler = (e: MediaQueryListEvent) => setIsMd(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, []);

  const handleSave = (items: { key: string; value: unknown }[]) => {
    updateSettings.mutate(items);
  };

  const ActiveComponent = TAB_COMPONENTS[activeTab];

  if (isLoading) {
    return (
      <>
        <Topbar title="설정" />
        <div className="flex items-center justify-center h-64 text-sm text-neutral-400">
          설정 불러오는 중...
        </div>
      </>
    );
  }

  return (
    <>
      <Topbar title="설정" />

      {isMd ? (
        /* ── PC: 좌측 탭 + 우측 폼 ── */
        <div className="flex min-h-[calc(100vh-56px)]">
          <nav className="w-48 shrink-0 border-r border-neutral-100 bg-neutral-50/50 py-3 px-2">
            {TABS.map((tab) => {
              const Icon = tab.icon;
              const active = activeTab === tab.key;
              return (
                <button
                  key={tab.key}
                  onClick={() => setActiveTab(tab.key)}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm font-medium transition mb-0.5
                    ${active
                      ? 'bg-neutral-900 text-white'
                      : 'text-neutral-500 hover:bg-neutral-100 hover:text-neutral-700'
                    }`}
                >
                  <Icon size={16} />
                  {tab.label}
                </button>
              );
            })}
          </nav>
          <div className="flex-1 px-6 py-5 max-w-2xl">
            <ActiveComponent
              settings={settings || {}}
              onSave={handleSave}
              saving={updateSettings.isPending}
            />
          </div>
        </div>
      ) : (
        /* ── 모바일: 아코디언 ── */
        <div className="px-4 py-3 space-y-1">
          {TABS.map((tab) => {
            const Icon = tab.icon;
            const isOpen = activeTab === tab.key;
            const Comp = TAB_COMPONENTS[tab.key];
            return (
              <div key={tab.key} className="border border-neutral-100 rounded-lg overflow-hidden">
                <button
                  onClick={() => setActiveTab(isOpen ? activeTab : tab.key)}
                  className="w-full flex items-center justify-between px-4 py-3 bg-neutral-50 text-sm font-medium"
                >
                  <span className="flex items-center gap-2">
                    <Icon size={16} className="text-neutral-500" />
                    {tab.label}
                  </span>
                  <ChevronDown
                    size={16}
                    className={`text-neutral-400 transition-transform ${isOpen ? 'rotate-180' : ''}`}
                  />
                </button>
                {isOpen && (
                  <div className="px-4 py-4 border-t border-neutral-100">
                    <Comp
                      settings={settings || {}}
                      onSave={handleSave}
                      saving={updateSettings.isPending}
                    />
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </>
  );
}