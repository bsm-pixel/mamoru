'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import { createClient } from '@/lib/supabase/client';
import { deliveryNet, deliveryOutstanding } from '@/lib/sales/amounts';

/** 납품 목록 */
export function useDeliveries(filters?: {
  status?: string;
  search?: string;
  dateRange?: string;
  page?: number;
  limit?: number;
}) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['deliveries', filters],
    queryFn: async () => {
      const params = new URLSearchParams();
      if (filters?.status) params.set('status', filters.status);
      if (filters?.search) params.set('search', filters.search);
      if (filters?.dateRange) params.set('date_range', filters.dateRange);
      if (filters?.page) params.set('page', String(filters.page));
      if (filters?.limit) params.set('limit', String(filters.limit));

      const res = await fetch(`/api/deliveries?${params.toString()}`);
      if (!res.ok) throw new Error('납품 목록 조회 실패');
      return res.json() as Promise<{ deliveries: Record<string, unknown>[]; total: number }>;
    },
  });
}

/** 납품 상세 */
export function useDelivery(id: string) {
  return useQuery({
    queryKey: ['delivery', id],
    queryFn: async () => {
      const res = await fetch(`/api/deliveries/${id}`);
      if (!res.ok) throw new Error('납품 상세 조회 실패');
      return res.json() as Promise<{
        delivery: Record<string, unknown>;
        items: Array<Record<string, unknown>>;
      }>;
    },
    enabled: !!id,
  });
}

/** 납품서 생성 */
export function useCreateDelivery() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      customer_id?: string;
      customer_name: string;
      customer_phone?: string;
      customer_type?: string;
      delivery_date?: string;
      expected_date?: string;
      memo?: string;
      vat_type?: string;
      receipt_type?: string;
      payment_status?: string;
      payment_method?: string;
      paid_amount?: number;
      discount_amount?: number;
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        category?: string;
        quantity: number;
        unit_price: number;
      }>;
    }) => {
      const res = await fetch('/api/deliveries', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : '생성 실패');
      }
      return res.json() as Promise<{ delivery: Record<string, unknown>; dlNumber: string }>;
    },
    onSuccess: (data) => {
      toast.success(`납품서 ${data.dlNumber} 생성 완료`);
      queryClient.invalidateQueries({ queryKey: ['deliveries'] });
      queryClient.invalidateQueries({ queryKey: ['delivery-stats'] });
    },
    onError: (err) => { toast.error('납품서 생성 실패: ' + String(err)); },
  });
}

/** 납품 상태 변경 / 편집 */
export function useUpdateDelivery() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...data }: { id: string; [key: string]: unknown }) => {
      const res = await fetch(`/api/deliveries/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : '업데이트 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['delivery', id] });
      queryClient.invalidateQueries({ queryKey: ['deliveries'] });
      queryClient.invalidateQueries({ queryKey: ['delivery-stats'] });
      queryClient.invalidateQueries({ queryKey: ['products'] });
    },
    onError: (err) => { toast.error('업데이트 실패: ' + String(err)); },
  });
}

/** 납품 탭 카운트 + 통계 */
export function useDeliveryStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['delivery-stats'],
    staleTime: 30_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const now = new Date();
      // KST 로컬 기준 — toISOString(UTC)는 오전 9시 이전 '이번주'를 하루 밀리게 함 (use-sales.ts 와 동일 패턴)
      const pad2 = (n: number) => String(n).padStart(2, '0');
      const monthStart = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-01`;
      const weekStartDate = new Date(now);
      const dow = now.getDay();
      weekStartDate.setDate(now.getDate() + (dow === 0 ? -6 : 1 - dow)); // Monday
      const weekStartStr = `${weekStartDate.getFullYear()}-${pad2(weekStartDate.getMonth() + 1)}-${pad2(weekStartDate.getDate())}`;

      const [allRes, draftRes, confirmedRes, shippedRes, settledRes, weekRes, monthRes, unpaidRes, monthRsRes] = await Promise.all([
        db.from('deliveries').select('*', { count: 'exact', head: true }).is('cancelled_at', null),
        db.from('deliveries').select('*', { count: 'exact', head: true }).eq('status', 'draft').is('cancelled_at', null),
        db.from('deliveries').select('*', { count: 'exact', head: true }).eq('status', 'confirmed').is('cancelled_at', null),
        db.from('deliveries').select('*', { count: 'exact', head: true }).eq('status', 'shipped').is('cancelled_at', null),
        db.from('deliveries').select('*', { count: 'exact', head: true }).eq('status', 'settled').is('cancelled_at', null),
        db.from('deliveries').select('total_amount, discount_amount').gte('delivery_date', weekStartStr).is('cancelled_at', null),
        db.from('deliveries').select('total_amount, discount_amount').gte('delivery_date', monthStart).is('cancelled_at', null),
        db.from('deliveries').select('total_amount, discount_amount, paid_amount').in('payment_status', ['unpaid', 'partial']).is('cancelled_at', null),
        // 이번달 복원수리(RS) 납품 항목 — 카드 '제품/복원수리' 분리표기용. monthRes 와 동일 필터로 조인.
        db.from('delivery_items').select('total_price, deliveries!inner(delivery_date, cancelled_at)').eq('category', 'RS').gt('total_price', 0).gte('deliveries.delivery_date', monthStart).is('deliveries.cancelled_at', null),
      ]);

      // 납품 total_amount는 이미 net(할인 반영) → deliveryNet/deliveryOutstanding 사용(할인 재차감 금지, SSOT)
      const sumAmount = (rows: Array<{ total_amount: number }>) =>
        rows.reduce((s, r) => s + deliveryNet(r), 0);

      const monthRepair = ((monthRsRes.data || []) as Array<{ total_price: number }>)
        .reduce((s, r) => s + (r.total_price || 0), 0);

      const unpaidAmount = (unpaidRes.data || []).reduce(
        (s: number, r: { total_amount: number; paid_amount?: number }) => s + deliveryOutstanding(r), 0
      );

      return {
        all: allRes.count || 0,
        draft: draftRes.count || 0,
        confirmed: confirmedRes.count || 0,
        shipped: shippedRes.count || 0,
        settled: settledRes.count || 0,
        weekAmount: sumAmount(weekRes.data || []),
        weekCount: (weekRes.data || []).length,
        monthAmount: sumAmount(monthRes.data || []),
        monthCount: (monthRes.data || []).length,
        monthRepair, // 이번달 납품 중 복원수리(RS) 합 — 제품 = monthAmount − monthRepair
        outstanding: unpaidAmount,
      };
    },
  });
}
