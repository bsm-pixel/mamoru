'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Repair, RepairInspection, RepairHistory } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 복원수리 목록 조회 */
export function useRepairs(filters?: {
  status?: string;
  search?: string;
  proceed_type?: string;
  proceed_type_neq?: string;
  page?: number;
  limit?: number;
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 20;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['repairs', filters],
    queryFn: async () => {
      let query = supabase
        .from('repairs')
        .select('*', { count: 'exact' })
        .order('received_at', { ascending: false })
        .range(from, to);

      if (filters?.status && filters.status !== 'all') {
        // 탭 그룹 필터
        const statusGroups: Record<string, string[]> = {
          intake: ['intake'],
          waiting: ['pickup_scheduled'],
          working: ['cost_notified', 'repairing'],
          shipping: ['ready_to_ship', 'shipped'],
          completed: ['delivered', 'completed'],
        };
        const group = statusGroups[filters.status];
        if (group) {
          query = query.in('status', group);
        } else {
          query = query.eq('status', filters.status);
        }
      }

      // proceed_type 필터
      if (filters?.proceed_type) {
        query = query.eq('proceed_type', filters.proceed_type);
      }
      if (filters?.proceed_type_neq) {
        query = query.neq('proceed_type', filters.proceed_type_neq);
      }

      if (filters?.search) {
        query = query.or(
          `name.ilike.%${filters.search}%,phone.ilike.%${filters.search}%,as_id.ilike.%${filters.search}%`
        );
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { repairs: (data || []) as Repair[], total: count || 0 };
    },
  });
}

/** 복원수리 단건 조회 (검수 + 이력 포함) */
export function useRepair(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['repair', id],
    queryFn: async () => {
      const [repairRes, inspectionsRes, historyRes] = await Promise.all([
        supabase.from('repairs').select('*').eq('id', id).single(),
        supabase
          .from('repair_inspections')
          .select('*')
          .eq('repair_id', id)
          .order('scissor_number', { ascending: true }),
        supabase
          .from('repair_history')
          .select('*')
          .eq('repair_id', id)
          .order('created_at', { ascending: false }),
      ]);

      if (repairRes.error) throw repairRes.error;
      return {
        repair: repairRes.data as Repair,
        inspections: (inspectionsRes.data || []) as RepairInspection[],
        history: (historyRes.data || []) as RepairHistory[],
      };
    },
    enabled: !!id,
  });
}

/** 상태 변경 */
export function useUpdateRepairStatus() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, status, note, ...rest }: {
      id: string;
      status: string;
      note?: string;
      [key: string]: unknown;
    }) => {
      const res = await fetch(`/api/repair/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status, note, ...rest }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '상태 변경 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('상태가 변경되었습니다');
      queryClient.invalidateQueries({ queryKey: ['repairs'] });
      queryClient.invalidateQueries({ queryKey: ['repair'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      queryClient.invalidateQueries({ queryKey: ['repair-dashboard-stats'] });
      queryClient.invalidateQueries({ queryKey: ['repair-tabs'] });
    },
    onError: (err) => {
      toast.error('상태 변경 실패: ' + String(err));
    },
  });
}

/** 필드 업데이트 (상태 변경 없이) */
export function useUpdateRepairFields() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...fields }: {
      id: string;
      [key: string]: unknown;
    }) => {
      const res = await fetch(`/api/repair/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(fields),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '수정 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('수정되었습니다');
      queryClient.invalidateQueries({ queryKey: ['repairs'] });
      queryClient.invalidateQueries({ queryKey: ['repair'] });
    },
    onError: (err) => {
      toast.error('수정 실패: ' + String(err));
    },
  });
}

/** 검수 데이터 저장 */
export function useSaveInspections() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ repairId, inspections }: {
      repairId: string;
      inspections: Partial<RepairInspection>[];
    }) => {
      const res = await fetch(`/api/repair/${repairId}/inspect`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ inspections }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '검수 저장 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('검수 데이터가 저장되었습니다');
      queryClient.invalidateQueries({ queryKey: ['repair'] });
    },
    onError: (err) => {
      toast.error('검수 저장 실패: ' + String(err));
    },
  });
}

/** 출고 (송장 생성) */
export function useShipRepair() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id }: { id: string }) => {
      const res = await fetch(`/api/repair/${id}/ship`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '출고 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`송장 생성 완료: ${data.invNo}`);
      queryClient.invalidateQueries({ queryKey: ['repairs'] });
      queryClient.invalidateQueries({ queryKey: ['repair'] });
    },
    onError: (err) => {
      toast.error('출고 실패: ' + String(err));
    },
  });
}

/** 송장 취소 */
export function useCancelShipment() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id }: { id: string }) => {
      const res = await fetch(`/api/repair/${id}/ship`, {
        method: 'DELETE',
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '송장 취소 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success('송장이 취소되었습니다');
      if (data?.warning) {
        toast(data.warning, { duration: 8000, icon: '⚠️' });
      }
      queryClient.invalidateQueries({ queryKey: ['repairs'] });
      queryClient.invalidateQueries({ queryKey: ['repair'] });
    },
    onError: (err) => {
      toast.error('송장 취소 실패: ' + String(err));
    },
  });
}

/** 알림톡 수동 발송 */
export function useSendRepairNotification() {
  return useMutation({
    mutationFn: async ({ repairId, template, extraData }: {
      repairId: string;
      template: string;
      extraData?: Record<string, string>;
    }) => {
      const res = await fetch(`/api/repair/${repairId}/notify`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ template, extraData }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '알림톡 발송 실패');
      }
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

/** 캐시 새로고침 */
export function useRepairSync() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      await queryClient.invalidateQueries({ queryKey: ['repairs'] });
      await queryClient.invalidateQueries({ queryKey: ['repair'] });
      await queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      await queryClient.invalidateQueries({ queryKey: ['repair-dashboard-stats'] });
      await queryClient.invalidateQueries({ queryKey: ['repair-tabs'] });
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
