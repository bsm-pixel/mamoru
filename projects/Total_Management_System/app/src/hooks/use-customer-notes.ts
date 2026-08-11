'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import { errMsg } from '@/lib/utils/err';

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
