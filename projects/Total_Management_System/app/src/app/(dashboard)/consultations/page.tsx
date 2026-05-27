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
import { AllConsultationsList } from '@/components/consultations/all-consultations-list';
import { ScheduleCalendar } from '@/components/consultations/schedule-calendar';
import { ConsultationDetailPanel } from '@/components/consultations/consultation-detail-panel';
import { SlidePanel } from '@/components/ui/slide-panel';
import { RefreshCw, Store, Truck, MessageCircle, Inbox, Loader, CheckCircle, MapPin, CalendarPlus, LayoutGrid, ArrowRight } from 'lucide-react';
import { CreateConsultationModal } from '@/components/consultations/create-consultation-modal';

// 카카오맵은 SSR 불가 → dynamic import
const FieldRequestMap = dynamic(
  () => import('@/components/consultations/field-request-map').then((m) => m.FieldRequestMap),
  { ssr: false }
);

type TabKey = 'all' | 'store_visit' | 'field_request' | 'talk_consult';

const TABS: { key: TabKey; label: string; icon: React.ReactNode }[] = [
  { key: 'all', label: '전체', icon: <LayoutGrid size={14} /> },
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

  const fieldCount = (fieldNew?.total || 0) + (fieldRe?.total || 0);
  const talkCount = talkNew?.total || 0;

  return {
    all: fieldCount + talkCount, // 전체: 출장 + 톡 합산
    store_visit: 0, // 매장은 확정만 → 대응필요 없음
    field_request: fieldCount,
    talk_consult: talkCount,
  };
}


// 서브탭 → 지도 activeStatuses 매핑
const SUB_TAB_STATUSES: Record<string, string[] | undefined> = {
  new_intake: ['pending_admin'],
  suggested: ['suggested'],
  action_needed: ['reschedule_requested', 'change_requested'],
  confirmed: ['confirmed'],
  past: ['completed'],  // 지난내역 탭에서만 완료 핀 노출 (cancelled은 위치 의미 X — 매핑 X)
};

export default function ConsultationsPage() {
  const [activeTab, setActiveTab] = useState<TabKey>('all');
  const [selectedFieldId, setSelectedFieldId] = useState<string | null>(null);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [fieldSubTab, setFieldSubTab] = useState<string>('new_intake');
  const [createModalOpen, setCreateModalOpen] = useState(false);
  // 077: 매장방문 서브탭 상태 — 상담완료 후 '확정' → '지난내역' 자동 전환용
  const [storeVisitTab, setStoreVisitTab] = useState<'confirmed' | 'past'>('confirmed');
  const sync = useConsultationSync();
  const needAction = useNeedActionCounts();
  const { data: stats } = useConsultationDashboardStats();

  // 077: 상담완료 후 매장방문 '확정' 탭에 있을 때만 '지난내역'으로 자동 전환
  const handleStoreVisitAfterComplete = () => {
    if (storeVisitTab === 'confirmed') setStoreVisitTab('past');
  };

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

      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 space-y-4">
        {/* 1행: 요약 3카드 (대시보드 CategoryCard 톤) + 새로고침 */}
        <div className="grid grid-cols-12 gap-3">
          <div className="col-span-12 lg:col-span-9 grid grid-cols-3 gap-3">
            <button
              onClick={() => setActiveTab('field_request')}
              className="bg-white rounded-2xl border border-stone-200 p-4 hover:border-stone-300 transition group text-left"
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <Inbox size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">신규</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className="text-3xl font-bold leading-none text-blue-600">{stats?.newIntake || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">확인 필요</p>
            </button>
            <button
              onClick={() => setActiveTab('field_request')}
              className="bg-white rounded-2xl border border-stone-200 p-4 hover:border-stone-300 transition group text-left"
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <Loader size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">진행</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className="text-3xl font-bold leading-none text-amber-600">{stats?.inProgress || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">일정 조율 중</p>
            </button>
            <button
              onClick={() => setActiveTab('store_visit')}
              className="bg-white rounded-2xl border border-stone-200 p-4 hover:border-stone-300 transition group text-left"
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-1.5">
                  <CheckCircle size={13} className="text-stone-400" />
                  <p className="text-[11px] text-stone-500 uppercase tracking-wider font-semibold">완료</p>
                </div>
                <ArrowRight size={12} className="text-stone-300 group-hover:text-stone-600 group-hover:translate-x-0.5 transition" />
              </div>
              <p className="text-3xl font-bold leading-none text-emerald-600">{stats?.completedMonth || 0}</p>
              <p className="text-[10px] text-stone-500 mt-1">이번달</p>
            </button>
          </div>
          <div className="col-span-12 lg:col-span-3 flex items-stretch">
            <button
              onClick={() => sync.mutate()}
              disabled={sync.isPending}
              className="w-full rounded-2xl border border-stone-200 bg-white hover:bg-stone-50 transition flex items-center justify-center gap-2 text-xs font-semibold text-stone-700 disabled:opacity-60 px-3 py-3"
            >
              <RefreshCw size={14} className={sync.isPending ? 'animate-spin' : ''} />
              {sync.isPending ? '새로고침 중...' : '새로고침'}
            </button>
          </div>
        </div>

        {/* 2행: 최상위 4탭 (전체+매장+출장+톡) + 일정 수동 등록 */}
        <div className="flex items-center justify-between border-b border-stone-200">
          <div className="flex gap-1">
            {TABS.map((tab) => (
              <button
                key={tab.key}
                onClick={() => { setActiveTab(tab.key); setSelectedFieldId(null); setSelectedId(null); }}
                className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 transition relative ${
                  activeTab === tab.key
                    ? 'border-stone-900 text-stone-900'
                    : 'border-transparent text-stone-500 hover:text-stone-700'
                }`}
              >
                {tab.icon}
                {tab.label}
                {needAction[tab.key] > 0 && (
                  <span className="ml-1 inline-flex items-center justify-center min-w-[18px] h-[18px] px-1 rounded-full bg-rose-500 text-white text-[10px] font-bold leading-none">
                    {needAction[tab.key]}
                  </span>
                )}
              </button>
            ))}
          </div>
          <button
            onClick={() => setCreateModalOpen(true)}
            className="px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition flex items-center gap-1.5 mb-1 shrink-0"
          >
            <CalendarPlus size={14} />
            일정수동등록
          </button>
        </div>

        {/* 메인 콘텐츠 — 탭별 분기 (전체 / 매장 / 출장 / 톡) */}
        {activeTab === 'all' ? (
          <>
            {/* 전체 — 모바일: 리스트 + 슬라이드 패널 */}
            {!isLg && (
              <div>
                <AllConsultationsList onSelect={setSelectedId} />
              </div>
            )}

            {/* 전체 — PC: 2열 (리스트 | 상세 모니터) */}
            {isLg && <div className="flex gap-4 h-[calc(100vh-260px)]">
              <div className="flex-1 min-w-0 overflow-y-auto">
                <AllConsultationsList onSelect={setSelectedId} />
              </div>
              <div className="w-[400px] shrink-0 overflow-y-auto">
                {selectedId ? (
                  <ConsultationDetailPanel consultationId={selectedId} />
                ) : (
                  <div className="flex flex-col items-center justify-center h-60 text-stone-400">
                    <LayoutGrid size={28} className="mb-2 opacity-40" />
                    <p className="text-xs text-center">목록에서 상담을<br />클릭하면 상세가 표시됩니다</p>
                  </div>
                )}
              </div>
            </div>}

            {/* 전체 — 모바일 슬라이드 패널 */}
            {!isLg && (
              <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="상담 상세">
                {selectedId && <ConsultationDetailPanel consultationId={selectedId} />}
              </SlidePanel>
            )}
          </>
        ) : activeTab === 'field_request' ? (
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
                <StoreVisitList onSelect={setSelectedId} tab={storeVisitTab} onTabChange={setStoreVisitTab} />
              </div>
            )}

            {/* 매장방문 — PC: 3열 (리스트 | 달력 | 상세모니터) */}
            {isLg && <div className="flex gap-4 h-[calc(100vh-220px)]">
              <div className="w-[400px] shrink-0 overflow-y-auto">
                <StoreVisitList onSelect={setSelectedId} tab={storeVisitTab} onTabChange={setStoreVisitTab} />
              </div>
              <div className="flex-1 min-w-0 overflow-y-auto">
                <ScheduleCalendar onSelect={setSelectedId} />
              </div>
              <div className="w-[400px] shrink-0 overflow-y-auto">
                {selectedId ? (
                  <ConsultationDetailPanel
                    consultationId={selectedId}
                    onAfterComplete={handleStoreVisitAfterComplete}
                  />
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
                {selectedId && (
                  <ConsultationDetailPanel
                    consultationId={selectedId}
                    onAfterComplete={handleStoreVisitAfterComplete}
                  />
                )}
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

      {/* 관리자 직접 상담 등록 모달 */}
      <CreateConsultationModal
        open={createModalOpen}
        onClose={() => setCreateModalOpen(false)}
        onCreated={(id) => {
          // 등록된 건을 상세 패널로 바로 열기
          setSelectedId(id);
        }}
      />
    </>
  );
}
