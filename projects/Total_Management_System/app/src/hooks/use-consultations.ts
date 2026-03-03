'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Consultation, ConsultationHistory } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 상담 목록 조회 (확장 필터 지원) */
export function useConsultations(filters?: {
  status?: string;
  statuses?: string[];        // 복수 상태 필터
  type?: string;
  search?: string;
  page?: number;
  limit?: number;
  dateFilter?: 'upcoming' | 'past' | 'all';  // visit_date 기준
  orderBy?: string;           // 'visit_date_asc' | 'updated_at_desc' 등
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 20;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['consultations', filters],
    queryFn: async () => {
      // 정렬 설정
      const orderCol = filters?.orderBy === 'visit_date_asc' ? 'visit_date'
        : filters?.orderBy === 'updated_at_desc' ? 'updated_at'
        : 'received_at';
      const ascending = filters?.orderBy === 'visit_date_asc';

      let query = supabase
        .from('consultations')
        .select('*', { count: 'exact' })
        .order(orderCol, { ascending })
        .range(from, to);

      // 단일 상태 필터
      if (filters?.status && filters.status !== 'all') {
        query = query.eq('status', filters.status);
      }
      // 복수 상태 필터
      if (filters?.statuses && filters.statuses.length > 0) {
        query = query.in('status', filters.statuses);
      }
      if (filters?.type && filters.type !== 'all') {
        query = query.eq('consultation_type', filters.type);
      }
      if (filters?.search) {
        query = query.or(
          `name.ilike.%${filters.search}%,phone.ilike.%${filters.search}%,unique_id.ilike.%${filters.search}%`
        );
      }
      // 날짜 필터
      const today = new Date().toISOString().slice(0, 10);
      if (filters?.dateFilter === 'upcoming') {
        query = query.gte('visit_date', today);
      } else if (filters?.dateFilter === 'past') {
        query = query.lt('visit_date', today);
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { consultations: (data || []) as Consultation[], total: count || 0 };
    },
  });
}

/** 상담 단건 조회 (이력 포함) */
export function useConsultation(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['consultation', id],
    queryFn: async () => {
      const [consultRes, historyRes] = await Promise.all([
        supabase.from('consultations').select('*').eq('id', id).single(),
        supabase
          .from('consultation_history')
          .select('*')
          .eq('consultation_id', id)
          .order('created_at', { ascending: false }),
      ]);

      if (consultRes.error) throw consultRes.error;
      return {
        consultation: consultRes.data as Consultation,
        history: (historyRes.data || []) as ConsultationHistory[],
      };
    },
    enabled: !!id,
  });
}

/** 상담 상태 변경 */
export function useUpdateConsultationStatus() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, status, note }: { id: string; status: string; note?: string }) => {
      const res = await fetch(`/api/consultation/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status, note }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('상태가 변경되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      queryClient.invalidateQueries({ queryKey: ['consultation-dashboard-stats'] });
    },
    onError: (err) => {
      toast.error('상태 변경 실패: ' + String(err));
    },
  });
}

/** 상담 동기화 — 캐시 새로고침 (GAS 자동 Push가 이미 Supabase에 저장함) */
export function useConsultationSync() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      // 모든 상담 관련 캐시를 강제 갱신
      await queryClient.invalidateQueries({ queryKey: ['consultations'] });
      await queryClient.invalidateQueries({ queryKey: ['consultation'] });
      await queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      await queryClient.invalidateQueries({ queryKey: ['consultation-dashboard-stats'] });
      return { synced: 0 };
    },
    onSuccess: () => {
      toast.success('새로고침 완료');
    },
    onError: (err) => {
      toast.error('새로고침 실패: ' + String(err));
    },
  });
}

/** 딜러 배정 */
export function useAssignDealer() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ consultationId, dealerId }: { consultationId: string; dealerId: string }) => {
      const res = await fetch('/api/consultation/assign', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ consultationId, dealerId }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('딜러가 배정되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
    },
    onError: (err) => {
      toast.error('딜러 배정 실패: ' + String(err));
    },
  });
}

// ============================================
// Phase 2-2: 추가 훅
// ============================================

/** 시간 제안 (출장요청) */
export function useSuggestTimes() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ consultationId, suggestions }: {
      consultationId: string;
      suggestions: { date: string; time: string }[];
    }) => {
      const res = await fetch('/api/consultation/suggest', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ consultationId, suggestions }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('시간 제안이 전송되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
    },
    onError: (err) => {
      toast.error('시간 제안 실패: ' + String(err));
    },
  });
}

/** 알림톡 발송 */
export function useSendNotification() {
  return useMutation({
    mutationFn: async ({ consultationId, template, extraData }: {
      consultationId: string;
      template: string;
      extraData?: Record<string, string>;
    }) => {
      const res = await fetch('/api/consultation/notify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ consultationId, template, extraData }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('알림톡이 발송되었습니다');
    },
    onError: (err) => {
      toast.error('알림톡 발송 실패: ' + String(err));
    },
  });
}

/** 보류 처리 (hold_reason 포함) */
export function useHoldConsultation() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, holdReason }: { id: string; holdReason: string }) => {
      const res = await fetch(`/api/consultation/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'on_hold', hold_reason: holdReason, note: `보류: ${holdReason}` }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('보류 처리되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
    },
    onError: (err) => {
      toast.error('보류 처리 실패: ' + String(err));
    },
  });
}

/** 출장 지연 안내 */
export function useFieldDelay() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ consultationId, delayMin }: {
      consultationId: string;
      delayMin: number;
    }) => {
      const res = await fetch('/api/consultation/delay', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ consultationId, delayMin }),
      });
      if (!res.ok) {
        const errBody = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(errBody.error || '지연 안내 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`지연 안내 발송 완료 (${data.delay_min}분, 도착 ${data.visit_time_revised})`);
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
    },
    onError: (err) => {
      toast.error('지연 안내 실패: ' + String(err));
    },
  });
}

/** 톡상담 시작 (상태변경 + talk_ready 알림톡) */
export function useStartTalkConsult() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id }: { id: string }) => {
      // 1) 상태를 in_progress로 변경
      const statusRes = await fetch(`/api/consultation/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'in_progress', note: '톡상담 시작' }),
      });
      if (!statusRes.ok) throw new Error(await statusRes.text());

      // 2) talk_ready 알림톡 발송
      await fetch('/api/consultation/notify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ consultationId: id, template: 'talk_ready' }),
      }).catch(() => {}); // 알림 실패해도 상태변경은 유지

      return statusRes.json();
    },
    onSuccess: () => {
      toast.success('톡상담 시작 + 알림톡 발송');
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      queryClient.invalidateQueries({ queryKey: ['consultation-dashboard-stats'] });
    },
    onError: (err) => {
      toast.error('톡상담 시작 실패: ' + String(err));
    },
  });
}

/** 일정 변경 (reschedule) */
export function useRescheduleConsultation() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, visitDate, visitTime, oldDate, oldTime, consultationType, uniqueId, notify }: {
      id: string;
      visitDate: string;
      visitTime: string;
      oldDate?: string;
      oldTime?: string;
      consultationType?: string; // 'store_visit' | 'field_request'
      uniqueId?: string;
      notify?: boolean;
    }) => {
      const res = await fetch(`/api/consultation/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          visit_date: visitDate,
          visit_time: visitTime,
          note: `일정 변경: ${visitDate} ${visitTime}`,
        }),
      });
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();

      // 알림톡 발송 (선택) — old/new 날짜 + change_request_link 포함
      if (notify) {
        const isField = consultationType === 'field_request';
        const template = isField ? 'field_rescheduled' : 'rescheduled';
        const changeLink = uniqueId
          ? `bsm-pixel.github.io/mamoru/projects/consulting/page_change_request.html?uid=${uniqueId}`
          : '';
        await fetch('/api/consultation/notify', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            consultationId: id,
            template,
            extraData: {
              old_date: oldDate || '',
              old_time: oldTime || '',
              new_date: visitDate,
              new_time: visitTime,
              change_request_link: changeLink,
            },
          }),
        }).catch(() => {}); // 실패해도 무시
      }

      return data;
    },
    onSuccess: () => {
      toast.success('일정이 변경되었습니다');
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['consultation'] });
    },
    onError: (err) => {
      toast.error('일정 변경 실패: ' + String(err));
    },
  });
}
