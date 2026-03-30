'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { Product } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 제품 단건 조회 */
export function useProduct(id: string) {
  return useQuery({
    queryKey: ['product', id],
    queryFn: async () => {
      const res = await fetch(`/api/products/${id}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ product: Product; supplier: { id: string; name: string } | null }>;
    },
    enabled: !!id,
  });
}

/** 제품 등록 */
export function useCreateProduct() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      sku: string;
      name: string;
      category: string;
      price: number;
      price_dealer?: number;
      price_academy?: number;
      price_purchase?: number;
      supplier_id?: string;
      description?: string;
      imweb_product_no?: string;
      barcode?: string;
      image_url?: string;
    }) => {
      const res = await fetch('/api/products', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ product: Product }>;
    },
    onSuccess: () => {
      toast.success('제품 등록 완료');
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => {
      toast.error('제품 등록 실패: ' + String(err));
    },
  });
}

/** 제품 수정 */
export function useUpdateProduct() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...data }: {
      id: string;
      name?: string;
      category?: string;
      price?: number;
      price_dealer?: number;
      price_academy?: number;
      price_purchase?: number;
      supplier_id?: string | null;
      description?: string | null;
      imweb_product_no?: string | null;
      barcode?: string | null;
      image_url?: string | null;
      is_active?: boolean;
      product_group?: string | null;
    }) => {
      const res = await fetch(`/api/products/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ product: Product }>;
    },
    onSuccess: (_, vars) => {
      toast.success('제품 수정 완료');
      queryClient.invalidateQueries({ queryKey: ['product', vars.id] });
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => {
      toast.error('수정 실패: ' + String(err));
    },
  });
}
