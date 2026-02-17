'use client';

import { useState } from 'react';
import dynamic from 'next/dynamic';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { useConsultationSync } from '@/hooks/use-consultations';
import { StoreVisitList } from '@/components/consultations/store-visit-list';
import { FieldRequestList } from '@/components/consultations/field-request-list';
import { TalkConsultList } from '@/components/consultations/talk-consult-list';
import { MobileFieldDayView } from '@/components/consultations/mobile-field-day-view';
import { ScheduleCalendar } from '@/components/consultations/schedule-calendar';
import { RefreshCw, Store, Truck, MessageCircle, Smartphone } from 'lucide-react';

// 카카오맵은 SSR 불가 → dynamic import
const FieldRequestMap = dynamic(
  () => import('@/components/consultations/field-request-map').then((m) => m.FieldRequestMap),
  { ssr: false }
);

type TabKey = 'store_visit' | 'field_request' | 'talk_consult';

const TABS: { key: TabKey; label: string; icon: React.ReactNode }[] = [
  { key: 'store_visit', label: '매장방문', icon: <Store size={14} /> },
  { key: 'field_request', label: '출장요청', icon: <Truck size={14} /> },
  { key: 'talk_consult', label: '톡상담', icon: <MessageCircle size={14} /> },
];

export default function ConsultationsPage() {
  const [activeTab, setActiveTab] = useState<TabKey>('store_visit');
  const [showMap, setShowMap] = useState(false);
  const [showMobileView, setShowMobileView] = useState(false);
  const sync = useConsultationSync();

  return (
    <>
      <Topbar title="상담관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 상단: 동기화 버튼 */}
        <div className="flex items-center justify-between">
          <Button
            variant="secondary"
            size="sm"
            onClick={() => sync.mutate()}
            disabled={sync.isPending}
          >
            <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
            {sync.isPending ? '동기화 중...' : 'GAS 동기화'}
          </Button>
        </div>

        {/* 최상위 3탭 */}
        <div className="flex gap-1 border-b border-neutral-200">
          {TABS.map((tab) => (
            <button
              key={tab.key}
              onClick={() => { setActiveTab(tab.key); setShowMap(false); }}
              className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 transition ${
                activeTab === tab.key
                  ? 'border-terracotta text-terracotta'
                  : 'border-transparent text-neutral-500 hover:text-neutral-700'
              }`}
            >
              {tab.icon}
              {tab.label}
            </button>
          ))}
        </div>

        {/* 메인 콘텐츠: 리스트 + 달력 사이드바 (PC) */}
        <div className="flex gap-6">
          {/* 좌측: 탭 콘텐츠 */}
          <div className="flex-1 min-w-0">
            {activeTab === 'store_visit' && <StoreVisitList />}
            {activeTab === 'field_request' && (
              <>
                <div className="flex justify-end mb-4">
                  <Button
                    variant={showMobileView ? 'primary' : 'ghost'}
                    size="sm"
                    onClick={() => setShowMobileView(!showMobileView)}
                  >
                    <Smartphone size={14} />
                    {showMobileView ? '리스트 보기' : '출장 일정'}
                  </Button>
                </div>

                {showMobileView ? (
                  <MobileFieldDayView />
                ) : (
                  <>
                    {showMap && <FieldRequestMap />}
                    <FieldRequestList
                      showMap={showMap}
                      onToggleMap={() => setShowMap(!showMap)}
                    />
                  </>
                )}
              </>
            )}
            {activeTab === 'talk_consult' && <TalkConsultList />}
          </div>

          {/* 우측: 달력 사이드바 (PC만, 매장방문/출장요청 탭일 때) */}
          {activeTab !== 'talk_consult' && (
            <div className="hidden lg:block w-[340px] shrink-0">
              <ScheduleCalendar />
            </div>
          )}
        </div>
      </div>
    </>
  );
}
