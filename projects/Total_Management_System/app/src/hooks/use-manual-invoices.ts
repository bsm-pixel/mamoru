'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import type { ManualInvoice } from '@/lib/supabase/types';

export interface CreateManualInvoiceInput {
  customer_id: string;
  goods_name: string;
  delivery_message?: string;
}

export interface CreateManualInvoiceResult {
  success: true;
  invoice?: ManualInvoice;
  warning?: string;
  invoiceNumber?: string;
}

/** 오늘 발급한 빠른 송장 (활성 건만, 최신순) */
export function useTodayManualInvoices() {
  return useQuery({
    queryKey: ['manual-invoices', 'today'],
    staleTime: 30_000,
    queryFn: async (): Promise<{ invoices: ManualInvoice[] }> => {
      const res = await fetch('/api/manual-invoices?date=today');
      if (!res.ok) return { invoices: [] };
      return res.json();
    },
  });
}

/** 특정 고객의 빠른 송장 이력 (취소 포함, 최신 20건) — 고객 상세 타임라인용 */
export function useCustomerManualInvoices(customerId: string | null | undefined) {
  return useQuery({
    queryKey: ['manual-invoices', 'customer', customerId],
    enabled: !!customerId,
    staleTime: 30_000,
    queryFn: async (): Promise<{ invoices: ManualInvoice[] }> => {
      const res = await fetch(`/api/manual-invoices?customer_id=${customerId}`);
      if (!res.ok) return { invoices: [] };
      return res.json();
    },
  });
}

/** 빠른 송장 발급 */
export function useCreateManualInvoice() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (input: CreateManualInvoiceInput): Promise<CreateManualInvoiceResult> => {
      const res = await fetch('/api/manual-invoices', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(input),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        // 주소·연락처 누락 차단은 customer_id를 함께 반환 — 호출자가 보강 링크 표시
        const err = new Error(typeof data.error === 'string' ? data.error : '송장 발급 실패') as Error & {
          customer_id?: string;
          missing?: 'address' | 'phone';
        };
        if (data.customer_id) err.customer_id = data.customer_id;
        if (data.missing) err.missing = data.missing;
        throw err;
      }
      return data;
    },
    onSuccess: (data) => {
      if (data.warning) toast(data.warning, { icon: '⚠️', duration: 6000 });
      else toast.success('송장 발급 완료');
      queryClient.invalidateQueries({ queryKey: ['manual-invoices'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 빠른 송장 취소 (soft delete + ALPS 취소 시도) */
export function useCancelManualInvoice() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, reason }: { id: string; reason?: string }): Promise<{ success: true; warning?: string }> => {
      const res = await fetch(`/api/manual-invoices/${id}`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ reason: reason ?? '' }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        throw new Error(typeof data.error === 'string' ? data.error : '취소 실패');
      }
      return data;
    },
    onSuccess: (data) => {
      if (data.warning) toast(data.warning, { icon: '⚠️', duration: 8000 });
      else toast.success('취소되었습니다');
      queryClient.invalidateQueries({ queryKey: ['manual-invoices'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}
