'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { PurchaseOrder, PurchaseOrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 발주 목록 */
export function usePurchaseOrders(filters?: {
  status?: string;
  search?: string;
  dateRange?: 'all' | 'today' | 'week' | 'month';
  page?: number;
  limit?: number;
}) {
  return useQuery({
    queryKey: ['purchase-orders', filters],
    queryFn: async () => {
      const params = new URLSearchParams();
      if (filters?.status) params.set('status', filters.status);
      if (filters?.search) params.set('search', filters.search);
      if (filters?.dateRange && filters.dateRange !== 'all') params.set('date_range', filters.dateRange);
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
      vat_type?: 'included' | 'separate' | 'none';
      currency?: string;
      exchange_rate?: number;
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
      expected_date?: string;
      supplier_name?: string;
      supplier_id?: string;
      order_date?: string;
      vat_type?: string;
      currency?: string;
      exchange_rate?: number;
      items?: Array<{ product_id?: string; product_name: string; sku?: string; quantity: number; unit_price: number }>;
      received_items?: Array<{ id: string; received_quantity: number }>; // 입고검수 실수령 수량
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

/** 취소된 발주 삭제 */
export function useDeletePurchaseOrder() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/purchasing/${id}`, { method: 'DELETE' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error));
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('발주 삭제 완료');
      queryClient.invalidateQueries({ queryKey: ['purchase-orders'] });
    },
    onError: (err) => toast.error('삭제 실패: ' + (err instanceof Error ? err.message : String(err))),
  });
}

// ══════════════════════════════════════════════════════════════
// 매입품목 카탈로그 (Supplier Product Catalog)
// ══════════════════════════════════════════════════════════════

export interface CatalogEntry {
  id: string;
  supplier_id: string;
  product_id: string;
  order_name: string;
  features: string;
  product_name: string;
  price_purchase: number;
  sku: string;
  category: string;
  product_group: string;
  sort_order: number;
}

/** 매입품목 조회 */
export function useSupplierCatalog(supplierId: string) {
  return useQuery({
    queryKey: ['supplier-catalog', supplierId],
    queryFn: async () => {
      const res = await fetch(`/api/suppliers/${supplierId}/catalog`);
      if (!res.ok) throw new Error('카탈로그 조회 실패');
      return res.json() as Promise<{ catalog: CatalogEntry[] }>;
    },
    enabled: !!supplierId,
  });
}

/** 매입품목 추가 */
export function useAddToCatalog() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ supplierId, productIds }: { supplierId: string; productIds: string[] }) => {
      const res = await fetch(`/api/suppliers/${supplierId}/catalog`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ product_ids: productIds }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, { supplierId }) => {
      toast.success('매입품목 추가 완료');
      queryClient.invalidateQueries({ queryKey: ['supplier-catalog', supplierId] });
    },
    onError: (err) => { toast.error('추가 실패: ' + String(err)); },
  });
}

/** 매입품목 수정 (주문명/특징) */
export function useUpdateCatalog() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ supplierId, catalogId, orderName, features }: {
      supplierId: string; catalogId: string; orderName?: string; features?: string;
    }) => {
      const res = await fetch(`/api/suppliers/${supplierId}/catalog`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ catalog_id: catalogId, order_name: orderName, features }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, { supplierId }) => {
      toast.success('수정 완료');
      queryClient.invalidateQueries({ queryKey: ['supplier-catalog', supplierId] });
    },
    onError: (err) => { toast.error('수정 실패: ' + String(err)); },
  });
}

/** 매입품목 삭제 */
export function useRemoveFromCatalog() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ supplierId, catalogId }: { supplierId: string; catalogId: string }) => {
      const res = await fetch(`/api/suppliers/${supplierId}/catalog`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ catalog_id: catalogId }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, { supplierId }) => {
      toast.success('매입품목 삭제 완료');
      queryClient.invalidateQueries({ queryKey: ['supplier-catalog', supplierId] });
    },
  });
}
