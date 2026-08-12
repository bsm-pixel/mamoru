'use client';

/**
 * 일정 — 상담(매장·출장) + 복원수리(직접방문·방문수거)를 한 화면에
 *
 * 시간축 전용 목적지 (IA v3, 2026-08-12 사장님 요청):
 *   1) 상단 요약 스트립(오늘/이번주/다음 일정)
 *   2) 통합 달력 + 선택일 타임라인 (공용 DashboardCalendarPanel 재사용)
 *   3) 다가오는 일정 리스트 (전화·길찾기 액션)
 *   색: 매장=초록 / 출장=보라 / 수리=주황 / 수거=파랑 (구글 캘린더 colorId도 동일 매칭)
 */

import { Topbar } from '@/components/layout/topbar';
import { DashboardCalendarPanel } from '@/components/dashboard/dashboard-calendar';
import { ScheduleSummary } from '@/components/schedule/schedule-summary';
import { UpcomingList } from '@/components/schedule/upcoming-list';

export default function SchedulePage() {
  return (
    <>
      <Topbar title="일정" />
      <div className="px-4 md:px-6 py-4 space-y-3">
        <ScheduleSummary />
        <DashboardCalendarPanel />
        <UpcomingList />
      </div>
    </>
  );
}
