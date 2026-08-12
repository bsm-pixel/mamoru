'use client';

/**
 * 일정 — 상담(매장·출장) + 복원수리 직접방문을 한 달력에 통합
 *
 * 시간축 전용 목적지 (IA v3, 2026-08-12 사장님 요청):
 *   - 상담관리·복원수리를 각각 눌러 확인하던 동선을 하나로
 *   - 기존 대시보드 통합 달력(DashboardCalendarPanel) 재사용 — 새 데이터 소스 없음
 *   - 색: 매장=초록 / 출장=보라 / 수리=주황 (구글 캘린더 colorId도 동일 매칭)
 */

import { Topbar } from '@/components/layout/topbar';
import { DashboardCalendarPanel } from '@/components/dashboard/dashboard-calendar';

export default function SchedulePage() {
  return (
    <>
      <Topbar title="일정" />
      <div className="px-4 md:px-6 py-4">
        <DashboardCalendarPanel />
      </div>
    </>
  );
}
