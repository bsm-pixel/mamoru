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
