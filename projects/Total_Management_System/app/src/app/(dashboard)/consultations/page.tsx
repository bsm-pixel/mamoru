'use client';

import { useState, useEffect } from 'react';
import dynamic from 'next/dynamic';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import { useConsultationSync, useConsultations } from '@/hooks/use-consultations';
import { useConsultationDashboardStats } from '@/hooks/use-dashboard-stats';
import { StoreVisitList } from '@/components/consultations/store-visit-list';
import { FieldRequestList } from '@/components/consultations/field-request-list';
import { TalkConsultList } from '@/components/consultations/talk-consult-list';
import { ScheduleCalendar } from '@/components/consultations/schedule-calendar';
import { ConsultationDetailPanel } from '@/components/consultations/consultation-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { RefreshCw, Store, Truck, MessageCircle, Inbox, Loader, CheckCircle, MapPin } from 'lucide-react';

// 카카오맵은 SSR 불가 → dynamic import
const FieldRequestMap = dynamic(
  () => import('@/components/consultations/field-request-map').then((m) => m.FieldRequestMap),
  { ssr: false }
);

type TabKey = 'store_visit' | 'field_request' | 'talk_consult';

const TABS: { key: TabKey; label: string; icon: React.ReactNode }[] = [
  { key: 'store_visit', label: '매장방문', icon: <Store size={14} /> },
  { key: 'field_request', label: '출장요청', icon: <Truck size={14} /> },
  { key: 'talk_consult', label: '온라인상담', icon: <MessageCircle size={14} /> },
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


// 서브탭 → 지도 activeStatuses 매핑
const SUB_TAB_STATUSES: Record<string, string[] | undefined> = {
  new_intake: ['pending_admin'],
  suggested: ['suggested'],
  action_needed: ['reschedule_requested', 'change_requested'],
  confirmed: ['confirmed'],
};

export default function ConsultationsPage() {
  const [activeTab, setActiveTab] = useState<TabKey>('store_visit');
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [fieldSubTab, setFieldSubTab] = useState<string>('new_intake');
  const sync = useConsultationSync();
  const needAction = useNeedActionCounts();
  const { data: stats } = useConsultationDashboardStats();

  // PC 여부 감지 (lg:1024px+) — SlidePanel 조건부 렌더링용
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
      <Topbar title="상담관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 상단: 요약 카드 + 동기화 버튼 */}
        <div className="flex items-center gap-3">
          <div className="flex gap-2 flex-1 min-w-0">
            <button
              onClick={() => setActiveTab('field_request')}
              className="flex items-center gap-2 px-3 py-2 rounded-lg bg-blue-50 hover:bg-blue-100 transition min-w-0"
            >
              <Inbox size={14} className="text-blue-600 shrink-0" />
              <span className="text-xs text-neutral-500">신규</span>
              <span className="text-sm font-bold text-blue-700">{stats?.newIntake || 0}</span>
            </button>
            <button
              onClick={() => setActiveTab('field_request')}
              className="flex items-center gap-2 px-3 py-2 rounded-lg bg-amber-50 hover:bg-amber-100 transition min-w-0"
            >
              <Loader size={14} className="text-amber-600 shrink-0" />
              <span className="text-xs text-neutral-500">진행</span>
              <span className="text-sm font-bold text-amber-700">{stats?.inProgress || 0}</span>
            </button>
            <button
              onClick={() => setActiveTab('store_visit')}
              className="flex items-center gap-2 px-3 py-2 rounded-lg bg-green-50 hover:bg-green-100 transition min-w-0"
            >
              <CheckCircle size={14} className="text-green-600 shrink-0" />
              <span className="text-xs text-neutral-500">완료</span>
              <span className="text-sm font-bold text-green-700">{stats?.completedMonth || 0}</span>
            </button>
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

        {/* 최상위 3탭 + #6 대응필요 뱃지 */}
        <div className="flex gap-1 border-b border-neutral-200">
          {TABS.map((tab) => (
            <button
              key={tab.key}
              onClick={() => { setActiveTab(tab.key); setSelectedFieldId(null); setSelectedId(null); }}
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

        {/* 메인 콘텐츠 — 출장 탭 PC: 3열 / 기타: 2열 */}
        {activeTab === 'field_request' ? (
          <>
            {/* 모바일: 기존 레이아웃 — PC에서는 렌더링하지 않음 */}
            {!isLg && (
              <div>
                <div className="mb-4">
                  <FieldRequestMap
                    selectedFieldId={selectedFieldId}
                    onFieldSelect={setSelectedFieldId}
                    onSelect={setSelectedId}
                    activeStatuses={SUB_TAB_STATUSES[fieldSubTab]}
                  />
                </div>
                <FieldRequestList
                  selectedFieldId={selectedFieldId}
                  onFieldSelect={setSelectedFieldId}
                  onSelect={setSelectedId}
                  onSubTabChange={setFieldSubTab}
                />
              </div>
            )}

            {/* PC: 3열 (리스트 | 지도 | 상세 모니터) */}
            {isLg && <div className="flex gap-4 h-[calc(100vh-220px)]">
              {/* 1열: 리스트 */}
              <div className="w-[480px] shrink-0 overflow-y-auto">
                <FieldRequestList
                  selectedFieldId={selectedFieldId}
                  onFieldSelect={setSelectedFieldId}
                  onSelect={setSelectedId}
                  onSubTabChange={setFieldSubTab}
                />
              </div>

              {/* 2열: 지도 */}
              <div className="flex-1 min-w-0 h-full">
                <FieldRequestMap
                  selectedFieldId={selectedFieldId}
                  onFieldSelect={setSelectedFieldId}
                  onSelect={setSelectedId}
                  activeStatuses={SUB_TAB_STATUSES[fieldSubTab]}
                />
              </div>

              {/* 3열: 상세 모니터 */}
              <div className="w-[400px] shrink-0 overflow-y-auto">
                {selectedId ? (
                  <ConsultationDetailPanel consultationId={selectedId} />
                ) : (
                  <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                    <MapPin size={28} className="mb-2 opacity-40" />
                    <p className="text-xs text-center">지도 마커 또는<br />리스트를 클릭하세요</p>
                  </div>
                )}
              </div>
            </div>}

            {/* 모바일 전용 슬라이드 패널 (PC에서는 3열 우측에 직접 표시) */}
            {!isLg && (
              <SlidePanel
                open={!!selectedId}
                onClose={() => setSelectedId(null)}
                title="상담 상세"
              >
                {selectedId && <ConsultationDetailPanel consultationId={selectedId} />}
              </SlidePanel>
            )}
          </>
        ) : activeTab === 'store_visit' ? (
          <>
            {/* 매장방문 — 모바일: 리스트 + 슬라이드 패널 */}
            {!isLg && (
              <div>
                <StoreVisitList onSelect={setSelectedId} />
              </div>
            )}

            {/* 매장방문 — PC: 3열 (리스트 | 달력 | 상세모니터) */}
            {isLg && <div className="flex gap-4 h-[calc(100vh-220px)]">
              <div className="w-[400px] shrink-0 overflow-y-auto">
                <StoreVisitList onSelect={setSelectedId} />
              </div>
              <div className="flex-1 min-w-0 overflow-y-auto">
                <ScheduleCalendar onSelect={setSelectedId} />
              </div>
              <div className="w-[400px] shrink-0 overflow-y-auto">
                {selectedId ? (
                  <ConsultationDetailPanel consultationId={selectedId} />
                ) : (
                  <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                    <Store size={28} className="mb-2 opacity-40" />
                    <p className="text-xs text-center">목록 또는 달력에서<br />일정을 클릭하세요</p>
                  </div>
                )}
              </div>
            </div>}

            {/* 매장방문 — 모바일 슬라이드 패널 */}
            {!isLg && (
              <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="상담 상세">
                {selectedId && <ConsultationDetailPanel consultationId={selectedId} />}
              </SlidePanel>
            )}
          </>
        ) : (
          /* 톡상담 — 기존 레이아웃 */
          <div>
            <TalkConsultList onSelect={setSelectedId} />
            <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="상담 상세">
              {selectedId && <ConsultationDetailPanel consultationId={selectedId} />}
            </SlidePanel>
          </div>
        )}
      </div>
    </>
  );
}
