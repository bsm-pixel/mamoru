'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Repair, RepairInspection, RepairHistory } from '@/lib/supabase/types';
import toast from 'react-hot-toast';
// 075: cross-domain invalidation — repair status/완료/출고/삭제 후 대시보드 매출 즉각 갱신
import { invalidateFinancialQueries } from '@/lib/query/invalidate-keys';

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
    staleTime: 30_000,
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

/** 상태 변경 (optimistic update) */
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
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '상태 변경 실패');
      }
      return res.json();
    },
    onMutate: async ({ id, status }) => {
      // 진행 중인 refetch 취소
      await queryClient.cancelQueries({ queryKey: ['repair', id] });
      await queryClient.cancelQueries({ queryKey: ['repair-tabs'] });

      // 스냅샷 저장
      const prevDetail = queryClient.getQueryData(['repair', id]);
      const prevTabs = queryClient.getQueryData(['repair-tabs']);

      // repair 상세 캐시 즉시 업데이트
      queryClient.setQueryData(['repair', id], (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const data = old as { repair: Repair; inspections: RepairInspection[]; history: RepairHistory[] };
        return { ...data, repair: { ...data.repair, status } };
      });

      return { prevDetail, prevTabs };
    },
    onError: (err, { id }, context) => {
      // 롤백
      if (context?.prevDetail) queryClient.setQueryData(['repair', id], context.prevDetail);
      if (context?.prevTabs) queryClient.setQueryData(['repair-tabs'], context.prevTabs);
      toast.error('상태 변경 실패: ' + String(err));
    },
    onSuccess: () => {
      toast.success('상태가 변경되었습니다');
    },
    onSettled: (_d, _e, { id }) => {
      // 서버 응답 후 캐시 재검증 + 대시보드 매출 갱신 (status 변경이 매출 카운트에 영향)
      queryClient.invalidateQueries({ queryKey: ['repair', id] });
      invalidateFinancialQueries(queryClient);
    },
  });
}

/** 필드 업데이트 (상태 변경 없이, optimistic update) */
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
        const msg = typeof err.error === 'string' ? err.error : JSON.stringify(err.error);
        throw new Error(msg || '수정 실패');
      }
      return res.json();
    },
    onMutate: async ({ id, ...fields }) => {
      await queryClient.cancelQueries({ queryKey: ['repair', id] });
      const prevDetail = queryClient.getQueryData(['repair', id]);

      queryClient.setQueryData(['repair', id], (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const data = old as { repair: Repair; inspections: RepairInspection[]; history: RepairHistory[] };
        return { ...data, repair: { ...data.repair, ...fields } };
      });

      return { prevDetail };
    },
    onError: (err, { id }, context) => {
      if (context?.prevDetail) queryClient.setQueryData(['repair', id], context.prevDetail);
      toast.error('수정 실패: ' + String(err));
    },
    onSuccess: () => {
      toast.success('수정되었습니다');
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['repair', id] });
      // 075: 필드 수정에 paid_at/total_amount/cost 등 매출 관련 필드가 포함될 수 있음 → 대시보드 즉각 갱신
      invalidateFinancialQueries(queryClient);
    },
  });
}

/** 복원수리 건 삭제 (알림톡 없이 완전 삭제) */
export function useDeleteRepair() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/repair/${id}`, { method: 'DELETE' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        const msg = typeof err.error === 'string' ? err.error : JSON.stringify(err.error);
        throw new Error(msg || '삭제 실패');
      }
      return res.json() as Promise<{ ok: boolean; deleted?: string; warning?: string }>;
    },
    onSuccess: async (data) => {
      // 접수 시점에 발송된 '새 복원수리 접수' OS 알림도 함께 회수
      if (data?.deleted && typeof navigator !== 'undefined' && 'serviceWorker' in navigator) {
        try {
          const reg = await navigator.serviceWorker.ready;
          reg.active?.postMessage({
            type: 'DISMISS',
            tag: `mamoru-as_received-${data.deleted}`,
          });
        } catch {
          /* SW 미등록/지원 안 됨 — 무시 */
        }
      }
      toast.success('삭제되었습니다');
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('삭제 실패: ' + String(err));
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
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '검수 저장 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('검수 데이터가 저장되었습니다');
      queryClient.invalidateQueries({ queryKey: ['repair'] });
      // 075: 검수 후 total_amount 자동 반영 → 대시보드 매출 즉각 갱신
      invalidateFinancialQueries(queryClient);
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
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '출고 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`송장 생성 완료: ${data.invNo}`);
      queryClient.invalidateQueries({ queryKey: ['repair'] });
      invalidateFinancialQueries(queryClient);
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
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '송장 취소 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success('송장이 취소되었습니다');
      if (data?.warning) {
        toast(data.warning, { duration: 8000, icon: '⚠️' });
      }
      queryClient.invalidateQueries({ queryKey: ['repair'] });
      invalidateFinancialQueries(queryClient);
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
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '알림톡 발송 실패');
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

/**
 * 복원수리 직접방문(당일수리) 일정 조회 — 달력 표시용 (2026-05-25 Phase 3-A)
 *   사장님 비전: TMS 달력에 매장방문/출장/직접방문 3종 일정을 함께 표시
 *
 * 조건: proceed_type='직접방문' + visit_date NOT NULL + status != 'cancelled'
 * useConsultations 와 동일 패턴 (createClient + staleTime 30초)
 */
export interface RepairScheduleItem {
  id: string;
  as_id: string;
  name: string;
  phone: string;
  visit_date: string;
  visit_time: string | null;
  visit_duration_min: number | null;
  status: string;
  qty_mamoru: number;
  qty_other: number;
}

export function useRepairSchedule(fromDate?: string, toDate?: string) {
  const supabase = createClient();
  return useQuery({
    queryKey: ['repair-schedule', fromDate, toDate],
    staleTime: 30_000,
    queryFn: async () => {
      let query = supabase
        .from('repairs')
        .select('id, as_id, name, phone, visit_date, visit_time, visit_duration_min, status, qty_mamoru, qty_other')
        .eq('proceed_type', '직접방문')
        .not('visit_date', 'is', null)
        .neq('status', 'cancelled');
      if (fromDate) query = query.gte('visit_date', fromDate);
      if (toDate) query = query.lte('visit_date', toDate);
      const { data, error } = await query;
      if (error) throw error;
      return (data || []) as RepairScheduleItem[];
    },
  });
}
