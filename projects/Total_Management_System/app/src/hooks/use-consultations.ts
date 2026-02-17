'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Consultation, ConsultationHistory } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 상담 목록 조회 */
export function useConsultations(filters?: {
  status?: string;
  type?: string;
  search?: string;
  page?: number;
  limit?: number;
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 20;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['consultations', filters],
    queryFn: async () => {
      let query = supabase
        .from('consultations')
        .select('*', { count: 'exact' })
        .order('received_at', { ascending: false })
        .range(from, to);

      if (filters?.status && filters.status !== 'all') {
        query = query.eq('status', filters.status);
      }
      if (filters?.type && filters.type !== 'all') {
        query = query.eq('consultation_type', filters.type);
      }
      if (filters?.search) {
        query = query.or(
          `name.ilike.%${filters.search}%,phone.ilike.%${filters.search}%,unique_id.ilike.%${filters.search}%`
        );
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
      queryClient.invalidateQueries({ queryKey: ['dashboard-stats'] });
    },
    onError: (err) => {
      toast.error('상태 변경 실패: ' + String(err));
    },
  });
}

/** 상담 동기화 (GAS → Supabase) */
export function useConsultationSync() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/consultation/sync', { method: 'POST' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`${data.synced}건 상담 동기화 완료`);
      queryClient.invalidateQueries({ queryKey: ['consultations'] });
      queryClient.invalidateQueries({ queryKey: ['dashboard-stats'] });
    },
    onError: (err) => {
      toast.error('상담 동기화 실패: ' + String(err));
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
