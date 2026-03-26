'use client';

import { useState } from 'react';
import dynamic from 'next/dynamic';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { useConsultationSync, useConsultations } from '@/hooks/use-consultations';
import { StoreVisitList } from '@/components/consultations/store-visit-list';
import { FieldRequestList } from '@/components/consultations/field-request-list';
import { TalkConsultList } from '@/components/consultations/talk-consult-list';
import { ScheduleCalendar } from '@/components/consultations/schedule-calendar';
import { ConsultationDetailPanel } from '@/components/consultations/consultation-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { RefreshCw, Store, Truck, MessageCircle } from 'lucide-react';

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

/** #6: 대응필요 상태 (탭별) */
function useNeedActionCounts() {
  // 출장: pending_admin + reschedule/change
  const { data: fieldNew } = useConsultations({ status: 'pending_admin', type: 'field_request', limit: 1 });
  const { data: fieldRe } = useConsultations({ statuses: ['reschedule_requested', 'change_requested'], type: 'field_request', limit: 1 });
  // 톡: pending_admin
  const { data: talkNew } = useConsultations({ status: 'pending_admin', type: 'talk_consult', limit: 1 });

  return {
    store_visit: 0, // 매장은 확정만 → 대응필요 없음
    field_request: (fieldNew?.total || 0) + (fieldRe?.total || 0),
    talk_consult: talkNew?.total || 0,
  };
}


export default function ConsultationsPage() {
  const [activeTab, setActiveTab] = useState<TabKey>('store_visit');
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const sync = useConsultationSync();
  const needAction = useNeedActionCounts();

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
            {sync.isPending ? '새로고침 중...' : '새로고침'}
          </Button>
        </div>

        {/* 최상위 3탭 + #6 대응필요 뱃지 */}
        <div className="flex gap-1 border-b border-neutral-200">
          {TABS.map((tab) => (
            <button
              key={tab.key}
              onClick={() => { setActiveTab(tab.key); setSelectedFieldId(null); }}
              className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 transition relative ${
                activeTab === tab.key
                  ? 'border-terracotta text-terracotta'
                  : 'border-transparent text-neutral-500 hover:text-neutral-700'
              }`}
            >
              {tab.icon}
              {tab.label}
              {needAction[tab.key] > 0 && (
                <span className="ml-1 inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-red-500 text-white text-[10px] font-bold leading-none">
                  {needAction[tab.key]}
                </span>
              )}
            </button>
          ))}
        </div>

        {/* R2: 메인 콘텐츠 — 탭별 사이드바 분기 */}
        <div className="flex gap-6">
          {/* 좌측: 탭 콘텐츠 */}
          <div className="flex-1 min-w-0">
            {activeTab === 'store_visit' && <StoreVisitList onSelect={setSelectedId} />}
            {activeTab === 'field_request' && (
              <>
                {/* 모바일: 리스트 상단 접이식 지도 */}
                <div className="lg:hidden mb-4">
                  <FieldRequestMap
                    selectedFieldId={selectedFieldId}
                    onFieldSelect={setSelectedFieldId}
                    onSelect={setSelectedId}
                  />
                </div>
                <FieldRequestList
                  selectedFieldId={selectedFieldId}
                  onFieldSelect={setSelectedFieldId}
                  onSelect={setSelectedId}
                />
              </>
            )}
            {activeTab === 'talk_consult' && <TalkConsultList onSelect={setSelectedId} />}
          </div>

          {/* R2: 우측 사이드바 — 출장 탭 = 지도 400px, 매장/톡 탭 = 달력 340px */}
          {activeTab === 'field_request' ? (
            <div className="hidden lg:block w-[400px] shrink-0">
              <div className="sticky top-16">
                <FieldRequestMap
                  selectedFieldId={selectedFieldId}
                  onFieldSelect={setSelectedFieldId}
                  onSelect={setSelectedId}
                />
              </div>
            </div>
          ) : activeTab !== 'talk_consult' ? (
            <div className="hidden lg:block w-[340px] shrink-0">
              <ScheduleCalendar />
            </div>
          ) : null}
        </div>

        {/* 상담 상세 슬라이드 패널 */}
        <SlidePanel
          open={!!selectedId}
          onClose={() => setSelectedId(null)}
          title="상담 상세"
        >
          {selectedId && <ConsultationDetailPanel consultationId={selectedId} />}
        </SlidePanel>
      </div>
    </>
  );
}
