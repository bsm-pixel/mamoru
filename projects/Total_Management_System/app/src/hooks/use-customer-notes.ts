'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import toast from 'react-hot-toast';
import { errMsg } from '@/lib/utils/err';

export interface LatestNote { customer_id: string; body: string; category: string | null; created_at: string; }

/** 여러 고객의 '최종 메모'를 한 번에 조회 → { customerId: 최근메모 } 맵 (목록 미리보기용) */
export function useLatestCustomerNotes(customerIds: (string | null | undefined)[]) {
  const ids = Array.from(new Set(customerIds.filter(Boolean))) as string[];
  return useQuery({
    queryKey: ['latest-customer-notes', [...ids].sort()],
    enabled: ids.length > 0,
    staleTime: 30_000,
    queryFn: async (): Promise<Record<string, LatestNote>> => {
      const sb = createClient();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = sb as any;
      const { data } = await db
        .from('customer_notes')
        .select('customer_id, body, category, created_at')
        .in('customer_id', ids)
        .is('deleted_at', null)
        .order('created_at', { ascending: false });
      const map: Record<string, LatestNote> = {};
      for (const n of (data || []) as LatestNote[]) {
        if (!map[n.customer_id]) map[n.customer_id] = n;   // 최신순이라 첫 번째가 최종
      }
      return map;
    },
  });
}

export interface CustomerNote {
  id: string;
  customer_id: string;
  body: string;
  category: string | null;
  created_at: string;
  created_by: string | null;
}

/** 고객 메모 타임라인 조회 (최신순) */
export function useCustomerNotes(customerId?: string | null) {
  return useQuery({
    queryKey: ['customer-notes', customerId],
    enabled: !!customerId,
    staleTime: 30_000,
    queryFn: async (): Promise<CustomerNote[]> => {
      const res = await fetch(`/api/customer-notes?customerId=${customerId}`);
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || '조회 실패');
      return (data.notes || []) as CustomerNote[];
    },
  });
}

/** 메모 추가 */
export function useAddCustomerNote() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ customerId, body, category }: { customerId: string; body: string; category?: string | null }) => {
      const res = await fetch('/api/customer-notes', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ customerId, body, category }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || '저장 실패');
      return data.note as CustomerNote;
    },
    onSuccess: (_d, v) => { qc.invalidateQueries({ queryKey: ['customer-notes', v.customerId] }); },
    onError: (e) => toast.error('메모 저장 실패: ' + errMsg(e)),
  });
}

/** 메모 삭제 (soft delete) */
export function useDeleteCustomerNote() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ id }: { id: string; customerId: string }) => {
      const res = await fetch(`/api/customer-notes?id=${id}`, { method: 'DELETE' });
      if (!res.ok) { const d = await res.json().catch(() => ({})); throw new Error(d.error || '삭제 실패'); }
      return true;
    },
    onSuccess: (_d, v) => { qc.invalidateQueries({ queryKey: ['customer-notes', v.customerId] }); },
    onError: (e) => toast.error('삭제 실패: ' + errMsg(e)),
  });
}
