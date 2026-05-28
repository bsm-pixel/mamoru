'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';

export interface BlackoutEntry {
  date: string;          // YYYY-MM-DD
  reason: string | null;
}

export interface BlackoutConsultation {
  id: string;
  visit_date: string;
  visit_time: string | null;
  consultation_type: 'store_visit' | 'field_request' | 'talk_consult';
  status: string;
  name: string;
  phone: string;
}

export interface BlackoutsData {
  blackouts: BlackoutEntry[];
  consultations: BlackoutConsultation[];
}

/** 휴무일 + 해당 기간 상담 일정 조회 (사장님 달력관리 전용) */
export function useBlackouts(from: string, to: string) {
  return useQuery<BlackoutsData>({
    queryKey: ['blackouts', from, to],
    staleTime: 30_000,
    queryFn: async () => {
      const res = await fetch(`/api/consultation/blackouts?from=${from}&to=${to}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
  });
}

/** 휴무일 추가/수정 (UPSERT) */
export function useCreateBlackout() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ date, reason }: { date: string; reason?: string }) => {
      const res = await fetch('/api/consultation/blackouts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date, reason }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '휴무일 등록 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('휴무일이 등록되었습니다');
      queryClient.invalidateQueries({ queryKey: ['blackouts'] });
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
    },
    onError: (err: Error) => toast.error('등록 실패: ' + err.message),
  });
}

/** 휴무일 해제 */
export function useDeleteBlackout() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (date: string) => {
      const res = await fetch(`/api/consultation/blackouts?date=${date}`, { method: 'DELETE' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '휴무일 해제 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('휴무일이 해제되었습니다');
      queryClient.invalidateQueries({ queryKey: ['blackouts'] });
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
    },
    onError: (err: Error) => toast.error('해제 실패: ' + err.message),
  });
}

// ============================================
// 정기 휴무 요일 (consultation_settings.disabled_weekdays)
// ============================================

export interface ConsultationSettings {
  start_hour: number;
  end_hour: number;
  duration_min: number;
  step_min: number;
  disabled_weekdays: number[];
  field_buffer_before: number;
  field_buffer_after: number;
}

/** consultation_settings 조회 */
export function useConsultationSettings() {
  return useQuery<ConsultationSettings>({
    queryKey: ['consultation-settings'],
    staleTime: 60_000,
    queryFn: async () => {
      const res = await fetch('/api/consultation/settings');
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
  });
}

/** 정기 휴무 요일 토글 (consultation_settings.disabled_weekdays UPSERT) */
export function useUpdateDisabledWeekdays() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (disabled_weekdays: number[]) => {
      const res = await fetch('/api/consultation/settings', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ disabled_weekdays }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '정기 휴무 변경 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('정기 휴무가 변경되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultation-settings'] });
      queryClient.invalidateQueries({ queryKey: ['blackouts'] });
    },
    onError: (err: Error) => toast.error('변경 실패: ' + err.message),
  });
}

/** 096: 영업 기본시간 업데이트 (start_hour/end_hour) — 달력관리 화면으로 이전 */
export function useUpdateBusinessHours() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ start_hour, end_hour }: { start_hour: number; end_hour: number }) => {
      const res = await fetch('/api/consultation/settings', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ start_hour, end_hour }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '영업시간 변경 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('영업시간이 변경되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultation-settings'] });
    },
    onError: (err: Error) => toast.error('변경 실패: ' + err.message),
  });
}

// ============================================
// 096: 날짜별 시간대 차단 (blocked_time_slots)
// ============================================

export interface BlockedSlot {
  id: string;
  date: string;        // YYYY-MM-DD
  start_time: string;  // HH:MM
  end_time: string;    // HH:MM
  reason: string | null;
}

/** 기간 시간대 차단 목록 조회 */
export function useBlockedSlots(from: string, to: string) {
  return useQuery<{ blockedSlots: BlockedSlot[] }>({
    queryKey: ['blocked-slots', from, to],
    staleTime: 30_000,
    queryFn: async () => {
      const res = await fetch(`/api/consultation/blocked-slots?from=${from}&to=${to}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
  });
}

/** 시간대 차단 추가 */
export function useCreateBlockedSlot() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ date, start_time, end_time, reason }: { date: string; start_time: string; end_time: string; reason?: string }) => {
      const res = await fetch('/api/consultation/blocked-slots', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ date, start_time, end_time, reason }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '시간대 차단 등록 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('시간대 차단이 등록되었습니다');
      queryClient.invalidateQueries({ queryKey: ['blocked-slots'] });
    },
    onError: (err: Error) => toast.error('등록 실패: ' + err.message),
  });
}

/** 시간대 차단 삭제 (id 단위) */
export function useDeleteBlockedSlot() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/consultation/blocked-slots?id=${id}`, { method: 'DELETE' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '시간대 차단 해제 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('시간대 차단이 해제되었습니다');
      queryClient.invalidateQueries({ queryKey: ['blocked-slots'] });
    },
    onError: (err: Error) => toast.error('해제 실패: ' + err.message),
  });
}
