'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';

/**
 * B2B 납품처(dealer/academy) 납품품목 카탈로그 hook
 * supplier catalog hook(`useSupplierCatalog`)과 mirror 패턴 (마이그 073)
 */

export interface CustomerCatalogEntry {
  id: string;
  customer_id: string;
  product_id: string;
  delivery_name: string;
  features: string;
  unit_price: number | null;  // 074: 거래처별 맞춤 단가
  sort_order: number;
  product_name: string;
  price: number;               // 제품 기본 소매가 (참고용)
  sku: string;
  category: string;
  product_group: string;
}

/** 납품품목 조회 */
export function useCustomerCatalog(customerId: string | null | undefined) {
  return useQuery({
    queryKey: ['customer-catalog', customerId],
    queryFn: async () => {
      const res = await fetch(`/api/customers/${customerId}/catalog`);
      if (!res.ok) throw new Error('카탈로그 조회 실패');
      return res.json() as Promise<{ catalog: CustomerCatalogEntry[] }>;
    },
    enabled: !!customerId,
  });
}

/** 납품품목 추가 (다중) */
export function useAddToCustomerCatalog() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ customerId, productIds }: { customerId: string; productIds: string[] }) => {
      const res = await fetch(`/api/customers/${customerId}/catalog`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ product_ids: productIds }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, { customerId }) => {
      toast.success('납품품목 추가 완료');
      queryClient.invalidateQueries({ queryKey: ['customer-catalog', customerId] });
    },
    onError: (err) => { toast.error('추가 실패: ' + String(err)); },
  });
}

/** 납품품목 수정 (납품명/특징) */
export function useUpdateCustomerCatalog() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ customerId, catalogId, deliveryName, features, unitPrice }: {
      customerId: string; catalogId: string; deliveryName?: string; features?: string; unitPrice?: number | null;
    }) => {
      const res = await fetch(`/api/customers/${customerId}/catalog`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ catalog_id: catalogId, delivery_name: deliveryName, features, unit_price: unitPrice }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, { customerId }) => {
      toast.success('수정 완료');
      queryClient.invalidateQueries({ queryKey: ['customer-catalog', customerId] });
    },
    onError: (err) => { toast.error('수정 실패: ' + String(err)); },
  });
}

/** 납품품목 삭제 (단건) */
export function useRemoveFromCustomerCatalog() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ customerId, catalogId }: { customerId: string; catalogId: string }) => {
      const res = await fetch(`/api/customers/${customerId}/catalog`, {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ catalog_id: catalogId }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, { customerId }) => {
      toast.success('납품품목 삭제 완료');
      queryClient.invalidateQueries({ queryKey: ['customer-catalog', customerId] });
    },
  });
}
