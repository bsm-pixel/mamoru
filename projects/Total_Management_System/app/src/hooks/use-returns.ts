'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { ReturnRow } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 반품·교환수거 목록 */
export function useReturns(opts?: { status?: string; search?: string }) {
  const params = new URLSearchParams();
  if (opts?.status) params.set('status', opts.status);
  if (opts?.search) params.set('search', opts.search);
  return useQuery({
    queryKey: ['returns', opts?.status, opts?.search],
    queryFn: async (): Promise<{ returns: ReturnRow[] }> => {
      const res = await fetch(`/api/returns?${params.toString()}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
  });
}

/** 반품 접수 생성 */
export function useCreateReturn() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (body: Record<string, unknown>): Promise<{ return: ReturnRow }> => {
      const res = await fetch('/api/returns', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => { qc.invalidateQueries({ queryKey: ['returns'] }); qc.invalidateQueries({ queryKey: ['schedule'] }); },
  });
}

/** 반품 상태 전이/수정 */
export function useUpdateReturn() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, ...body }: { id: string } & Record<string, unknown>): Promise<{ return: ReturnRow }> => {
      const res = await fetch(`/api/returns/${id}`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('반품 상태가 변경되었습니다');
      qc.invalidateQueries({ queryKey: ['returns'] });
      qc.invalidateQueries({ queryKey: ['schedule'] });
      qc.invalidateQueries({ queryKey: ['inventory'] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : '변경 실패'),
  });
}

/** 교환 출고 송장 발행 (새 제품 1개 롯데 송장) */
export function useShipReturn() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (id: string): Promise<{ return: ReturnRow }> => {
      const res = await fetch(`/api/returns/${id}/ship`, { method: 'POST' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('교환 출고 송장이 발행되었습니다');
      qc.invalidateQueries({ queryKey: ['returns'] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : '송장 발행 실패'),
  });
}

/** 캘린더용 — 방문/택배 수거 예약(pickup_date) 있는 진행중 반품 */
export interface ReturnPickupItem {
  id: string;
  return_number: string;
  name: string | null;
  phone: string | null;
  pickup_date: string;
  pickup_method: string | null;
  status: string;
  address: string | null;
}
export function useReturnPickupSchedule(monthStart: string, monthEnd: string) {
  return useQuery({
    queryKey: ['return-pickup-schedule', monthStart, monthEnd],
    queryFn: async (): Promise<ReturnPickupItem[]> => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = createClient() as any;
      const { data } = await db.from('returns')
        .select('id, return_number, name, phone, pickup_date, pickup_method, status, address')
        .not('pickup_date', 'is', null)
        .gte('pickup_date', monthStart)
        .lte('pickup_date', monthEnd)
        .not('status', 'in', '(completed,cancelled)');
      return data || [];
    },
  });
}
