'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { ProductSerial, Product } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 제품별 시리얼 목록 */
export function useSerials(productId?: string, filters?: {
  status?: string;
  search?: string;
  page?: number;
  limit?: number;
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 50;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['serials', productId, filters],
    staleTime: 30_000,
    queryFn: async () => {
      let query = supabase
        .from('product_serials')
        .select('*', { count: 'exact' })
        .order('created_at', { ascending: false })
        .range(from, to);

      if (productId) {
        query = query.eq('product_id', productId);
      }
      if (filters?.status && filters.status !== 'all') {
        query = query.eq('status', filters.status);
      }
      if (filters?.search) {
        query = query.or(
          `serial_number.ilike.%${filters.search}%,barcode.ilike.%${filters.search}%,sold_to_name.ilike.%${filters.search}%`
        );
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { serials: (data || []) as ProductSerial[], total: count || 0 };
    },
  });
}

/** 특정 제품의 판매 가능(in_stock + ready) 시리얼 목록 — 판매 시 선택용 */
export function useAvailableSerials(productId: string | undefined) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['serials-available', productId],
    queryFn: async () => {
      const { data, error } = await supabase
        .from('product_serials')
        .select('id, serial_number, barcode, warehouse_zone, manufactured_at, memo')
        .eq('product_id', productId!)
        .eq('status', 'in_stock')
        .eq('warehouse_zone', 'ready')  // ready만 판매 가능
        .order('serial_number');
      if (error) throw error;
      return (data || []) as Array<{
        id: string;
        serial_number: string;
        barcode: string | null;
        warehouse_zone: string;
        manufactured_at: string | null;
        memo: string | null;
      }>;
    },
    enabled: !!productId,
    staleTime: 10 * 1000,
  });
}

/** 바코드/시리얼로 단건 조회 */
export function useSerialLookup() {
  const supabase = createClient();

  return useMutation({
    mutationFn: async (code: string) => {
      // 시리얼넘버로 검색
      let { data } = await supabase
        .from('product_serials')
        .select('*')
        .eq('serial_number', code)
        .single();

      if (!data) {
        // 바코드로 검색
        const res = await supabase
          .from('product_serials')
          .select('*')
          .eq('barcode', code)
          .single();
        data = res.data;
      }

      if (!data) throw new Error('시리얼/바코드를 찾을 수 없습니다');
      return data as ProductSerial;
    },
  });
}

/** 시리얼 등록 */
export function useCreateSerial() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      product_id: string;
      serial_number: string;
      barcode?: string;
      lot_number?: string;
      manufactured_at?: string;
      memo?: string;
    }) => {
      const res = await fetch('/api/serials', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('시리얼 등록 완료');
      queryClient.invalidateQueries({ queryKey: ['serials'] });
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => {
      toast.error('시리얼 등록 실패: ' + String(err));
    },
  });
}

/** 시리얼 일괄 등록 */
export function useCreateSerialBatch() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      product_id: string;
      count: number;
      prefix?: string;
      lot_number?: string;
    }) => {
      const res = await fetch('/api/serials/batch', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`${data.created}개 시리얼 생성 완료`);
      queryClient.invalidateQueries({ queryKey: ['serials'] });
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => {
      toast.error('일괄 등록 실패: ' + String(err));
    },
  });
}

/** 시리얼 일괄 zone 변경 */
export function useUpdateSerialZone() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: { ids: string[]; warehouse_zone: 'raw' | 'ready' | 'display' }) => {
      const res = await fetch('/api/serials', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_data, vars) => {
      const zoneLabel = { raw: '매입원본', ready: '판매준비', display: '진열' }[vars.warehouse_zone];
      toast.success(`${vars.ids.length}개 → ${zoneLabel}으로 이동`);
      queryClient.invalidateQueries({ queryKey: ['serials'] });
      queryClient.invalidateQueries({ queryKey: ['serials-available'] });
      queryClient.invalidateQueries({ queryKey: ['inventory'] });
    },
    onError: (err) => {
      toast.error('zone 변경 실패: ' + String(err));
    },
  });
}

/** 시리얼 상태 변경 */
export function useUpdateSerialStatus() {
  const queryClient = useQueryClient();
  const supabase = createClient();

  return useMutation({
    mutationFn: async (data: { id: string; status: string; memo?: string }) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { error } = await (supabase as any)
        .from('product_serials')
        .update({ status: data.status, memo: data.memo })
        .eq('id', data.id);
      if (error) throw error;
    },
    onSuccess: () => {
      toast.success('상태 변경 완료');
      queryClient.invalidateQueries({ queryKey: ['serials'] });
    },
    onError: (err) => {
      toast.error('상태 변경 실패: ' + String(err));
    },
  });
}
