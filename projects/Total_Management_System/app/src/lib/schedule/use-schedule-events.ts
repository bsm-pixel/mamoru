'use client';

/**
 * useScheduleEvents — 일정 달력 공용 데이터 소스 (DRY, 2026-08-12)
 *   상담(매장·출장) + 복원수리(직접방문·방문수거)를 월 범위로 조회해
 *   날짜별(YYYY-MM-DD) 이벤트 맵으로 반환.
 *   대시보드·일정 페이지·상담 매장탭 달력이 이 단일 훅을 공유(기존 복붙 3벌 통합).
 */

import { useMemo } from 'react';
import { useConsultations } from '@/hooks/use-consultations';
import {
  useRepairSchedule,
  useRepairPickupSchedule,
  type RepairScheduleItem,
  type RepairPickupItem,
} from '@/hooks/use-repairs';
import { activityDisplay } from '@/lib/customer/display';
import type { ScheduleCategory } from './colors';
import type { Consultation } from '@/lib/supabase/types';

export type ScheduleEvent =
  | { kind: 'consult'; data: Consultation }
  | { kind: 'repair'; data: RepairScheduleItem }   // 복원수리 직접방문
  | { kind: 'pickup'; data: RepairPickupItem };    // 복원수리 방문수거(수거예정)

/** 정렬 키 — 시간 있는 건 시간순, 수거(시간 없음)는 맨 앞 */
function eventTime(ev: ScheduleEvent): string {
  if (ev.kind === 'consult') return ev.data.visit_time || '';
  if (ev.kind === 'repair') return ev.data.visit_time || '';
  return ''; // pickup
}

export function useScheduleEvents(monthStart: string, monthEnd: string) {
  // 매장방문(확정/진행)
  const { data: storeData } = useConsultations({
    statuses: ['confirmed', 'in_progress'],
    type: 'store_visit',
    limit: 200,
    dateFilter: 'all',
    orderBy: 'visit_date_asc',
  });
  // 출장(확정/제안/진행)
  const { data: fieldData } = useConsultations({
    statuses: ['confirmed', 'suggested', 'in_progress'],
    type: 'field_request',
    limit: 200,
    dateFilter: 'all',
    orderBy: 'visit_date_asc',
  });
  // 복원수리 직접방문 / 방문수거 (훅 자체가 월 범위로 필터)
  const { data: repairData } = useRepairSchedule(monthStart, monthEnd);
  const { data: pickupData } = useRepairPickupSchedule(monthStart, monthEnd);

  const dateMap = useMemo(() => {
    const map = new Map<string, ScheduleEvent[]>();
    const add = (date: string, ev: ScheduleEvent) => {
      if (!map.has(date)) map.set(date, []);
      map.get(date)!.push(ev);
    };
    for (const c of storeData?.consultations || []) {
      if (!c.visit_date || c.visit_date < monthStart || c.visit_date > monthEnd) continue;
      add(c.visit_date, { kind: 'consult', data: c });
    }
    for (const c of fieldData?.consultations || []) {
      if (!c.visit_date || c.visit_date < monthStart || c.visit_date > monthEnd) continue;
      add(c.visit_date, { kind: 'consult', data: c });
    }
    for (const r of repairData || []) {
      if (!r.visit_date) continue;
      add(r.visit_date, { kind: 'repair', data: r });
    }
    for (const p of pickupData || []) {
      if (!p.pickup_date) continue;
      add(p.pickup_date, { kind: 'pickup', data: p });
    }
    for (const arr of map.values()) arr.sort((a, b) => eventTime(a).localeCompare(eventTime(b)));
    return map;
  }, [storeData, fieldData, repairData, pickupData, monthStart, monthEnd]);

  return { dateMap };
}

// ── 이벤트 필드 접근 공용 헬퍼 (요약·다가오는 일정 등에서 재사용) ──

export function eventCategory(ev: ScheduleEvent): ScheduleCategory {
  if (ev.kind === 'consult') return ev.data.consultation_type === 'store_visit' ? 'store' : 'field';
  if (ev.kind === 'repair') return 'repair_visit';
  return 'repair_pickup';
}

export function eventName(ev: ScheduleEvent): string {
  if (ev.kind === 'consult') return activityDisplay(ev.data.activity_name, ev.data.name);
  return ev.data.name || '고객';
}

export function eventPhone(ev: ScheduleEvent): string {
  return ev.data.phone || '';
}

export function eventTimeStr(ev: ScheduleEvent): string | null {
  return eventTime(ev) || null;
}

export function eventHref(ev: ScheduleEvent): string {
  return ev.kind === 'consult' ? `/consultations/${ev.data.id}` : `/repairs/${ev.data.id}`;
}

/** 출장 상담만 방문 주소 반환 (길찾기용). 나머지는 null */
export function eventAddress(ev: ScheduleEvent): string | null {
  if (ev.kind === 'consult' && ev.data.consultation_type === 'field_request') {
    return [ev.data.address_road, ev.data.address_detail].filter(Boolean).join(' ') || null;
  }
  return null;
}

/** dateMap → 날짜(오름차순, 같은 날은 시간순) 평탄 배열 */
export function flattenEvents(dateMap: Map<string, ScheduleEvent[]>): { date: string; ev: ScheduleEvent }[] {
  const out: { date: string; ev: ScheduleEvent }[] = [];
  for (const [date, evs] of dateMap) for (const ev of evs) out.push({ date, ev });
  // 날짜 오름차순 (같은 날짜는 hook이 이미 시간순 → stable sort로 유지)
  out.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  return out;
}
