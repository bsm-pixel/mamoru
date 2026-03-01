'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { PurchaseOrder, PurchaseOrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 발주 목록 */
export function usePurchaseOrders(filters?: {
  status?: string;
  search?: string;
  page?: number;
  limit?: number;
}) {
  return useQuery({
    queryKey: ['purchase-orders', filters],
    queryFn: async () => {
      const params = new URLSearchParams();
      if (filters?.status) params.set('status', filters.status);
      if (filters?.search) params.set('search', filters.search);
      if (filters?.page) params.set('page', String(filters.page));
      if (filters?.limit) params.set('limit', String(filters.limit));

      const res = await fetch(`/api/purchasing?${params}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ orders: PurchaseOrder[]; total: number }>;
    },
  });
}

/** 발주 상세 */
export function usePurchaseOrder(id: string) {
  return useQuery({
    queryKey: ['purchase-order', id],
    queryFn: async () => {
      const res = await fetch(`/api/purchasing/${id}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        order: PurchaseOrder;
        items: PurchaseOrderItem[];
        supplier: { id: string; name: string } | null;
      }>;
    },
    enabled: !!id,
  });
}

/** 발주 생성 */
export function useCreatePurchaseOrder() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      supplier_id?: string;
      supplier_name: string;
      order_date?: string;
      expected_date?: string;
      memo?: string;
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
      }>;
    }) => {
      const res = await fetch('/api/purchasing', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ order: PurchaseOrder; poNumber: string }>;
    },
    onSuccess: (data) => {
      toast.success(`발주 생성: ${data.poNumber}`);
      queryClient.invalidateQueries({ queryKey: ['purchase-orders'] });
    },
    onError: (err) => {
      toast.error('발주 생성 실패: ' + String(err));
    },
  });
}

/** 발주 상태 전환 */
export function useUpdatePurchaseOrder() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...data }: {
      id: string;
      status?: string;
      deposit_amount?: number;
      balance_amount?: number;
      received_date?: string;
      memo?: string;
    }) => {
      const res = await fetch(`/api/purchasing/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ order: PurchaseOrder }>;
    },
    onSuccess: (_, vars) => {
      toast.success('발주 업데이트 완료');
      queryClient.invalidateQueries({ queryKey: ['purchase-order', vars.id] });
      queryClient.invalidateQueries({ queryKey: ['purchase-orders'] });
    },
    onError: (err) => {
      toast.error('업데이트 실패: ' + String(err));
    },
  });
}
