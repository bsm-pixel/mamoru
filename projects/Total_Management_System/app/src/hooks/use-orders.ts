'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Order, OrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 주문 목록 조회 */
export function useOrders(filters?: {
  status?: string;
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
    queryKey: ['orders', filters],
    queryFn: async () => {
      let query = supabase
        .from('orders')
        .select('*', { count: 'exact' })
        .order('ordered_at', { ascending: false })
        .range(from, to);

      if (filters?.status && filters.status !== 'all') {
        query = query.eq('status', filters.status);
      }
      if (filters?.search) {
        query = query.or(
          `orderer_name.ilike.%${filters.search}%,recipient_name.ilike.%${filters.search}%,imweb_order_no.ilike.%${filters.search}%,invoice_number.ilike.%${filters.search}%`
        );
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { orders: (data || []) as Order[], total: count || 0 };
    },
  });
}

/** 상태별 주문 건수 조회 (탭 배지용) */
export function useOrderCounts() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['order-counts'],
    queryFn: async () => {
      const statuses = ['pay_done', 'preparing', 'shipping', 'delivered', 'cancel_pending', 'cancelled'];
      const counts: Record<string, number> = {};

      const results = await Promise.all(
        statuses.map(async (s) => {
          const { count } = await supabase
            .from('orders')
            .select('*', { count: 'exact', head: true })
            .eq('status', s);
          return { status: s, count: count || 0 };
        })
      );

      let total = 0;
      results.forEach((r) => {
        counts[r.status] = r.count;
        total += r.count;
      });
      counts['all'] = total;

      return counts;
    },
    staleTime: 30_000, /* 30초 캐시 — 탭 전환마다 재요청 방지 */
  });
}

/** 주문 단건 조회 (품목 포함) */
export function useOrder(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['order', id],
    queryFn: async () => {
      const [orderRes, itemsRes] = await Promise.all([
        supabase.from('orders').select('*').eq('id', id).single(),
        supabase.from('order_items').select('*').eq('order_id', id),
      ]);

      if (orderRes.error) throw orderRes.error;
      return {
        order: orderRes.data as Order,
        items: (itemsRes.data || []) as OrderItem[],
      };
    },
    enabled: !!id,
  });
}

/** 주문 동기화 */
export function useOrderSync() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/imweb/sync', { method: 'POST' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`${data.synced}건 동기화 완료`);
      queryClient.invalidateQueries({ queryKey: ['orders'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      queryClient.invalidateQueries({ queryKey: ['order-dashboard-stats'] });
    },
    onError: (err) => {
      toast.error('동기화 실패: ' + String(err));
    },
  });
}

/** 송장 생성 */
export function useBookInvoice() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      orderId: string;
      ordNo: string;
      rcvName: string;
      rcvTel: string;
      rcvZip: string;
      rcvAdr: string;
      gdsNm?: string;
      dlvMsg?: string;
      ordSct?: '1' | '2' | '3';
    }) => {
      const res = await fetch('/api/lotte/book', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`송장 생성 완료: ${data.invNo}`);
      if (data.imwebNeedsManual) {
        toast('아임웹에서 "배송대기 처리" 후 자동 연동됩니다', {
          icon: '⚠️',
          duration: 6000,
        });
      }
      queryClient.invalidateQueries({ queryKey: ['orders'] });
      queryClient.invalidateQueries({ queryKey: ['order'] });
      queryClient.invalidateQueries({ queryKey: ['hub-stats'] });
      queryClient.invalidateQueries({ queryKey: ['order-dashboard-stats'] });
    },
    onError: (err) => {
      toast.error('송장 생성 실패: ' + String(err));
    },
  });
}

/** 송장 취소 */
export function useCancelInvoice() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: { invNo: string; orderId: string }) => {
      const res = await fetch('/api/lotte/cancel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('ALPS에서 직접 집하취소 해주세요', { duration: 5000 });
      queryClient.invalidateQueries({ queryKey: ['orders'] });
      queryClient.invalidateQueries({ queryKey: ['order'] });
    },
    onError: (err) => {
      toast.error('취소 처리 실패: ' + String(err));
    },
  });
}
