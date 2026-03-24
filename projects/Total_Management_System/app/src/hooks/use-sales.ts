'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { OfflineSale, OfflineSaleItem, Product } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 오프라인 판매 목록 */
export function useSales(filters?: {
  search?: string;
  page?: number;
  limit?: number;
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 20;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['sales', filters],
    staleTime: 30_000,
    queryFn: async () => {
      let query = supabase
        .from('offline_sales')
        .select('*', { count: 'exact' })
        .order('sale_date', { ascending: false })
        .range(from, to);

      if (filters?.search) {
        query = query.or(
          `customer_name.ilike.%${filters.search}%,customer_phone.ilike.%${filters.search}%,sale_number.ilike.%${filters.search}%`
        );
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { sales: (data || []) as OfflineSale[], total: count || 0 };
    },
  });
}

/** 오프라인 판매 단건 조회 (항목 + 시리얼 포함) */
export function useSale(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sale', id],
    queryFn: async () => {
      const [saleRes, itemsRes, serialsRes] = await Promise.all([
        supabase.from('offline_sales').select('*').eq('id', id).single(),
        supabase.from('offline_sale_items').select('*').eq('sale_id', id),
        supabase.from('product_serials').select('id, serial_number, product_id').eq('offline_sale_id', id),
      ]);
      if (saleRes.error) throw saleRes.error;
      const sale = saleRes.data as OfflineSale;

      // customer_id가 있으면 최신 연락처 가져오기
      if (sale.customer_id) {
        const { data: cust } = await supabase
          .from('customers')
          .select('phone')
          .eq('id', sale.customer_id)
          .single();
        if (cust?.phone) {
          (sale as Record<string, unknown>).customer_phone = cust.phone;
        }
      }

      return {
        sale,
        items: (itemsRes.data || []) as OfflineSaleItem[],
        serials: (serialsRes.data || []) as Array<{ id: string; serial_number: string; product_id: string }>,
      };
    },
    enabled: !!id,
  });
}

/** 제품 목록 (판매 입력용) */
export function useProducts() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['products'],
    staleTime: 5 * 60 * 1000,
    queryFn: async () => {
      const { data, error } = await supabase
        .from('products')
        .select('*')
        .eq('is_active', true)
        .order('category')
        .order('name');
      if (error) throw error;
      return (data || []) as Product[];
    },
  });
}

/** 오프라인 판매 생성 */
export function useCreateSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      sale: {
        customer_id?: string;
        customer_name: string;
        customer_phone?: string;
        sale_date?: string;
        total_amount: number;
        discount_amount?: number;
        paid_amount: number;
        payment_method: string;
        payment_status?: string;
        memo?: string;
        sale_channel?: string;
      };
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        serial_ids?: string[];
      }>;
    }) => {
      const res = await fetch('/api/sales', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`판매 등록 완료: ${data.saleNumber}`);
      queryClient.invalidateQueries({ queryKey: ['sales'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
    },
    onError: (err) => {
      toast.error('판매 등록 실패: ' + String(err));
    },
  });
}

/** 판매 취소 — 서버 확인 필수 (시리얼/재고/아임웹 역전) */
export function useCancelSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, reason }: { id: string; reason?: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'cancel', reason }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '판매 취소 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('판매가 취소되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sales'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => {
      toast.error('판매 취소 실패: ' + String(err));
    },
  });
}

/** 결제상태 변경 — 낙관적 업데이트 */
export function useUpdatePaymentStatus() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, payment_status, paid_amount }: {
      id: string;
      payment_status: string;
      paid_amount?: number;
    }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'update_payment', payment_status, paid_amount }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '결제상태 변경 실패');
      }
      return res.json();
    },
    onMutate: async ({ id, payment_status, paid_amount }) => {
      await queryClient.cancelQueries({ queryKey: ['sale', id] });

      const prevDetail = queryClient.getQueryData(['sale', id]);

      // 상세 캐시 즉시 업데이트
      queryClient.setQueryData(['sale', id], (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const data = old as { sale: OfflineSale; items: OfflineSaleItem[]; serials: unknown[] };
        return {
          ...data,
          sale: {
            ...data.sale,
            payment_status,
            ...(paid_amount !== undefined ? { paid_amount } : {}),
          },
        };
      });

      return { prevDetail };
    },
    onError: (err, { id }, context) => {
      if (context?.prevDetail) queryClient.setQueryData(['sale', id], context.prevDetail);
      toast.error('결제상태 변경 실패: ' + String(err));
    },
    onSuccess: () => {
      toast.success('결제상태가 변경되었습니다');
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['sales'] });
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
    },
  });
}

/** 메모 수정 — 낙관적 업데이트 */
export function useUpdateSaleMemo() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, memo }: { id: string; memo: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'update_memo', memo }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '메모 수정 실패');
      }
      return res.json();
    },
    onMutate: async ({ id, memo }) => {
      await queryClient.cancelQueries({ queryKey: ['sale', id] });
      const prevDetail = queryClient.getQueryData(['sale', id]);

      queryClient.setQueryData(['sale', id], (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const data = old as { sale: OfflineSale; items: OfflineSaleItem[]; serials: unknown[] };
        return { ...data, sale: { ...data.sale, memo } };
      });

      return { prevDetail };
    },
    onError: (err, { id }, context) => {
      if (context?.prevDetail) queryClient.setQueryData(['sale', id], context.prevDetail);
      toast.error('메모 수정 실패: ' + String(err));
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
    },
  });
}

