'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { Product } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** PATCH /api/products/[id] 의 아임웹 상품명 반영 결과 */
export interface ImwebNameResult {
  status: 'synced' | 'failed' | 'skipped';
  error?: string;
}

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

/** 연결된 소싱 정보 역조회 (제품/부자재 상세에서 업체·링크·단가 표시) */
export interface LinkedSourcingInfo {
  sticker_no: string;
  supplier_name: string | null;
  supplier_url: string | null;
  vendor_url: string | null;
  unit_price: number;
  features_memo: string | null;
  krw_price: number;
  exchange_rate: number;
  po_number: string | null;
}
export function useSourcingByProduct(productId: string) {
  return useQuery({
    queryKey: ['sourcing-by-product', productId],
    queryFn: async () => {
      const res = await fetch(`/api/sourcing/by-product?product_id=${productId}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ sourcing: LinkedSourcingInfo | null }>;
    },
    enabled: !!productId,
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
      price_groups?: Record<string, { price?: number; display_name?: string }>;
      supplier_id?: string;
      description?: string;
      imweb_product_no?: string;
      barcode?: string;
      image_url?: string;
      purchase_name?: string;
      purchase_url?: string;
      tags?: Record<string, unknown>;
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
      price_groups?: Record<string, { price?: number; display_name?: string }>;
      dealer_name?: string | null;
      academy_name?: string | null;
      stock_quantity?: number;
      supplier_id?: string | null;
      description?: string | null;
      imweb_product_no?: string | null;
      barcode?: string | null;
      image_url?: string | null;
      is_active?: boolean;
      product_group?: string | null;
      purchase_name?: string | null;
      tags?: Record<string, unknown> | null;
      /** 이름이 같아도 아임웹 상품명을 TMS 이름으로 다시 맞춤 */
      sync_imweb_name?: boolean;
    }) => {
      const res = await fetch(`/api/products/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) {
        const errBody = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof errBody.error === 'string' ? errBody.error : JSON.stringify(errBody.error));
      }
      return res.json() as Promise<{ product: Product; imwebName?: ImwebNameResult }>;
    },
    onSuccess: (res, vars) => {
      // 아임웹 상품명 반영 결과 (2026-11-09) — 실패는 TMS 저장과 별개로 분명히 알린다
      if (res.imwebName?.status === 'failed') {
        toast.error(`TMS에는 저장됐지만 아임웹 상품명은 못 바꿨습니다 — ${res.imwebName.error}\n제품 상세의 「아임웹 이름 맞추기」로 다시 시도하세요`, { duration: 8000 });
      } else if (res.imwebName?.status === 'synced') {
        toast.success('제품 수정 완료 · 아임웹 상품명도 바꿨습니다');
      } else {
        toast.success('제품 수정 완료');
      }
      queryClient.invalidateQueries({ queryKey: ['product', vars.id] });
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => {
      toast.error('수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}
