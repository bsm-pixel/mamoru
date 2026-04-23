'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Order, OrderItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 주문 목록 조회 */
export function useOrders(filters?: {
  status?: string;
  search?: string;
  dateRange?: 'all' | 'today' | 'week' | 'month';
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
    staleTime: 30_000,
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
      if (filters?.dateRange && filters.dateRange !== 'all') {
        const now = new Date();
        let dateFrom: string;
        if (filters.dateRange === 'today') dateFrom = now.toISOString().slice(0, 10);
        else if (filters.dateRange === 'week') { const d = new Date(now); d.setDate(d.getDate() - 7); dateFrom = d.toISOString().slice(0, 10); }
        else { const d = new Date(now); d.setMonth(d.getMonth() - 1); dateFrom = d.toISOString().slice(0, 10); }
        query = query.gte('ordered_at', dateFrom);
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { orders: (data || []) as Order[], total: count || 0 };
    },
  });
}

/** 상태별 주문 건수 조회 (탭 배지용) — RPC 1회 호출 */
export function useOrderCounts() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['order-counts'],
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any).rpc('get_order_counts');

      if (!error && data) {
        const d = data as Record<string, number>;
        const total = Object.values(d).reduce((s, v) => s + (v || 0), 0);
        return { ...d, all: total } as Record<string, number>;
      }

      // Fallback: RPC 미배포 시 개별 쿼리
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
      results.forEach((r) => { counts[r.status] = r.count; total += r.count; });
      counts['all'] = total;
      return counts;
    },
    staleTime: 60_000, // 60초 — RPC 통합 후 여유있게
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
      queryClient.invalidateQueries({ queryKey: ['order-counts'] });
      // hub-stats / order-dashboard-stats → staleTime 자연 갱신에 위임
    },
    onError: (err) => {
      toast.error('동기화 실패: ' + String(err));
    },
  });
}

/** 아임웹 상품 동기화 결과 (Phase 2: 투명한 실패 집계) */
export interface ProductSyncResult {
  success: boolean;
  total_fetched: number; // 아임웹 API에서 받은 전체 개수
  synced: number;        // DB 반영 성공 개수
  created: number;
  updated: number;
  linked: number;        // 수동 등록된 상품에 imweb_product_no 연결
  errors: string[];
}

/** 아임웹 상품 동기화 */
export function useProductSync() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/imweb/sync-products', { method: 'POST' });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<ProductSyncResult>;
    },
    onSuccess: (data) => {
      const failed = data.errors?.length || 0;
      if (failed > 0) {
        toast(
          `상품 ${data.synced}/${data.total_fetched}건 반영 (${failed}건 실패)`,
          { icon: '⚠️', duration: 6000 }
        );
      } else {
        toast.success(`상품 ${data.synced}건 동기화 완료 (생성 ${data.created} · 업데이트 ${data.updated} · 연결 ${data.linked})`);
      }
      queryClient.invalidateQueries({ queryKey: ['products'] });
      queryClient.invalidateQueries({ queryKey: ['inventory'] });
      queryClient.invalidateQueries({ queryKey: ['sync-logs'] });
    },
    onError: (err) => {
      toast.error('상품 동기화 실패: ' + String(err));
    },
  });
}

/** 송장 생성 — 낙관적 업데이트 */
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
    onMutate: async (data) => {
      await queryClient.cancelQueries({ queryKey: ['orders'] });
      await queryClient.cancelQueries({ queryKey: ['order', data.orderId] });

      // 스냅샷 저장
      const prevOrders = queryClient.getQueriesData({ queryKey: ['orders'] });
      const prevOrder = queryClient.getQueryData(['order', data.orderId]);

      // 목록 캐시에서 해당 주문 즉시 '배송중' 표시
      queryClient.setQueriesData({ queryKey: ['orders'] }, (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const d = old as { orders: Order[]; total: number };
        return {
          ...d,
          orders: d.orders.map((o) =>
            o.id === data.orderId ? { ...o, status: 'shipping' as const, invoice_number: '생성중...' } : o
          ),
        };
      });

      return { prevOrders, prevOrder };
    },
    onError: (err, data, context) => {
      // 롤백
      if (context?.prevOrders) {
        for (const [key, value] of context.prevOrders) {
          queryClient.setQueryData(key, value);
        }
      }
      if (context?.prevOrder) {
        queryClient.setQueryData(['order', data.orderId], context.prevOrder);
      }
      toast.error('송장 생성 실패: ' + String(err));
    },
    onSuccess: (data) => {
      toast.success(`송장 생성 완료: ${data.invNo}`);
      if (data.imwebNeedsManual) {
        toast('아임웹에서 "배송대기 처리" 후 자동 연동됩니다', {
          icon: '⚠️',
          duration: 6000,
        });
      }
    },
    onSettled: (_d, _e, data) => {
      queryClient.invalidateQueries({ queryKey: ['orders'] });
      queryClient.invalidateQueries({ queryKey: ['order', data.orderId] });
      queryClient.invalidateQueries({ queryKey: ['order-counts'] });
    },
  });
}

/** 주문 직접 취소 (송장 없는 주문: pay_done → cancelled) */
export function useCancelOrder() {
  const supabase = createClient();
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (orderId: string) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const { error } = await db
        .from('orders')
        .update({ status: 'cancelled' })
        .eq('id', orderId);
      if (error) throw error;
    },
    onSuccess: (_d, orderId) => {
      toast.success('주문이 취소되었습니다');
      queryClient.invalidateQueries({ queryKey: ['orders'] });
      queryClient.invalidateQueries({ queryKey: ['order', orderId] });
      queryClient.invalidateQueries({ queryKey: ['order-counts'] });
    },
    onError: (err) => {
      toast.error('취소 실패: ' + String(err));
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
    onSuccess: (_d, data) => {
      toast.success('ALPS에서 직접 집하취소 해주세요', { duration: 5000 });
      queryClient.invalidateQueries({ queryKey: ['orders'] });
      queryClient.invalidateQueries({ queryKey: ['order', data.orderId] });
      queryClient.invalidateQueries({ queryKey: ['order-counts'] });
    },
    onError: (err) => {
      toast.error('취소 처리 실패: ' + String(err));
    },
  });
}
