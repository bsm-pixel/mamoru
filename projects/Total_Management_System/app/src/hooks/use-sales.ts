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
      return {
        sale: saleRes.data as OfflineSale,
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

