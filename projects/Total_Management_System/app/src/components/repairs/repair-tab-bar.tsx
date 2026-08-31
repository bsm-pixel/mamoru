'use client';

export type RepairTabKey =
  | 'intake'
  | 'pickup_needed'
  | 'visit_scheduled'
  | 'inbound_waiting'
  | 'in_progress'
  | 'ready_to_ship'
  | 'shipped'
  | 'recall';   // 재수리 — 재수거 접수했으나 재작업 전 (출고완료에 묻히던 건)

export interface RepairTabDef {
  key: RepairTabKey;
  label: string;
  count: number;
}

interface RepairTabBarProps {
  tabs: RepairTabDef[];
  activeTab: RepairTabKey;
  onTabChange: (tab: RepairTabKey) => void;
}

/** 고정 탭 바 — 수평 스크롤, 카운트 뱃지 */
export function RepairTabBar({ tabs, activeTab, onTabChange }: RepairTabBarProps) {
  return (
    <div className="sticky top-0 z-10 bg-warm-ivory border-b border-neutral-200">
      <div className="flex overflow-x-auto scrollbar-hide">
        {tabs.map((tab) => {
          const isActive = activeTab === tab.key;
          return (
            <button
              key={tab.key}
              onClick={() => onTabChange(tab.key)}
              className={`shrink-0 flex items-center gap-1.5 px-4 py-3 text-sm font-semibold border-b-2 transition whitespace-nowrap ${
                isActive
                  ? 'border-terracotta text-terracotta'
                  : 'border-transparent text-neutral-500 hover:text-neutral-700'
              }`}
            >
              {tab.label}
              {tab.count > 0 && (
                <span
                  className={`inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 rounded-full text-[11px] font-bold ${
                    isActive
                      ? 'bg-terracotta text-white'
                      : 'bg-neutral-200 text-neutral-600'
                  }`}
                >
                  {tab.count}
                </span>
              )}
            </button>
          );
        })}
      </div>
    </div>
  );
}
