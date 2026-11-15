'use client';

/**
 * 일정 — 상담(매장·출장) + 복원수리(직접방문·방문수거)를 한 화면에
 *
 * 시간축 전용 목적지 (IA v3, 2026-08-12 사장님 요청):
 *   1) 상단 요약 스트립(오늘/이번주/다음 일정)
 *   2) 통합 달력 + 선택일 타임라인 (공용 DashboardCalendarPanel 재사용)
 *   3) 다가오는 일정 리스트 (전화·길찾기 액션)
 *   색: 매장=초록 / 출장=보라 / 수리=주황 / 수거=파랑 (구글 캘린더 colorId도 동일 매칭)
 *
 * 2026-11-15 일정을 넣는 버튼 추가(사장님 지시) — 전엔 상담관리 · 복원수리 화면으로 가야만 넣을 수 있었다.
 *   새 창을 만들지 않고 그 두 화면의 등록 창을 그대로 쓴다(같은 규칙 · 같은 알림톡).
 */

import { useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import { CalendarPlus, Wrench } from 'lucide-react';
import { Topbar } from '@/components/layout/topbar';
import { DashboardCalendarPanel } from '@/components/dashboard/dashboard-calendar';
import { ScheduleSummary } from '@/components/schedule/schedule-summary';
import { UpcomingList } from '@/components/schedule/upcoming-list';
import { CreateConsultationModal } from '@/components/consultations/create-consultation-modal';
import { VisitBookingModal } from '@/components/repairs/visit-booking-modal';

export default function SchedulePage() {
  const queryClient = useQueryClient();
  const [consultOpen, setConsultOpen] = useState(false);
  const [repairOpen, setRepairOpen] = useState(false);
  // 등록하면 달력 · 요약 · 다가오는 일정이 바로 바뀌게 관련 자료를 모두 다시 받는다
  const refresh = () => queryClient.invalidateQueries();

  return (
    <>
      <Topbar title="일정" />
      <div className="px-4 md:px-6 py-4 space-y-3">
        <div className="flex items-center gap-2 flex-wrap">
          <button type="button" data-new-action onClick={() => setConsultOpen(true)}
            className="h-9 px-3 rounded-lg bg-stone-900 text-white text-xs font-semibold hover:bg-stone-800 transition inline-flex items-center gap-1.5">
            <CalendarPlus size={14} />상담 일정 등록
          </button>
          <button type="button" onClick={() => setRepairOpen(true)}
            className="h-9 px-3 rounded-lg border border-stone-300 bg-white text-stone-700 text-xs font-semibold hover:bg-stone-50 transition inline-flex items-center gap-1.5">
            <Wrench size={14} />복원수리 방문 예약
          </button>
          <span className="text-[11px] text-neutral-400">고객이 접수한 건은 자동으로 들어옵니다</span>
        </div>
        <ScheduleSummary />
        <DashboardCalendarPanel includePast />
        <UpcomingList />
      </div>

      <CreateConsultationModal
        open={consultOpen}
        onClose={() => setConsultOpen(false)}
        onCreated={() => refresh()}
      />
      <VisitBookingModal
        open={repairOpen}
        onClose={() => setRepairOpen(false)}
        onCreated={() => refresh()}
      />
    </>
  );
}
