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
import { RefreshCw, Store, Truck, MessageCircle, CalendarCheck, ChevronDown } from 'lucide-react';

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

/** #2: 오늘 일정 건수 */
function useTodayCount() {
  const today = new Date().toISOString().slice(0, 10);
  const { data } = useConsultations({
    status: 'confirmed',
    limit: 1,
    // visit_date = today는 API에서 dateFilter로 처리 안 되므로 전체 confirmed에서 클라이언트 필터
  });
  // 간단히: confirmed 중 visit_date가 오늘인 건수를 별도 쿼리
  const { data: todayData } = useConsultations({
    statuses: ['confirmed', 'assigned', 'in_progress'],
    limit: 1,
  });
  // 실제로는 서버에서 날짜 필터가 필요하지만, 가벼운 카운트 쿼리로 대체
  return todayData?.consultations?.filter(c => c.visit_date === today).length || 0;
}

export default function ConsultationsPage() {
  const [activeTab, setActiveTab] = useState<TabKey>('store_visit');
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [todayBarOpen, setTodayBarOpen] = useState(false);
  const sync = useConsultationSync();
  const needAction = useNeedActionCounts();

  // #2: 오늘 일정 (모바일용)
  const today = new Date().toISOString().slice(0, 10);
  const { data: todayData } = useConsultations({
    statuses: ['confirmed', 'assigned', 'in_progress'],
    limit: 20,
  });
  const todaySchedule = todayData?.consultations?.filter(c => c.visit_date === today) || [];

  return (
    <>
      <Topbar title="상담관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* #2: 모바일 오늘 일정 요약 바 */}
        {todaySchedule.length > 0 && (
          <div className="lg:hidden">
            <button
              onClick={() => setTodayBarOpen(!todayBarOpen)}
              className="w-full flex items-center justify-between px-3 py-2.5 rounded-lg bg-terracotta/5 border border-terracotta/20 transition"
            >
              <div className="flex items-center gap-2">
                <CalendarCheck size={16} className="text-terracotta" />
                <span className="text-sm font-semibold text-terracotta">오늘 {todaySchedule.length}건</span>
              </div>
              <ChevronDown size={16} className={`text-terracotta transition-transform ${todayBarOpen ? 'rotate-180' : ''}`} />
            </button>
            {todayBarOpen && (
              <div className="mt-1 rounded-lg border border-neutral-200 bg-white divide-y divide-neutral-100 overflow-hidden">
                {todaySchedule.map(c => (
                  <button
                    key={c.id}
                    onClick={() => setSelectedId(c.id)}
                    className="w-full flex items-center justify-between px-3 py-2 text-left hover:bg-warm-ivory/60 transition"
                  >
                    <span className="text-sm font-medium truncate">{c.name}</span>
                    <span className="text-xs text-neutral-500 shrink-0 ml-2">{c.visit_time || '시간미정'}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

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
