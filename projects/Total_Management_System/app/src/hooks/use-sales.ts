'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { OfflineSale, OfflineSaleItem, Product } from '@/lib/supabase/types';
import toast from 'react-hot-toast';
// 075: cross-domain invalidation 일원화 — mutation 후 대시보드/통계가 즉각 갱신되도록
import { invalidateFinancialQueries } from '@/lib/query/invalidate-keys';

/** 판매 탭 타입 */
export type SalesTab = 'all' | 'today' | 'unpaid' | 'cancelled';
export type SalesChannel = 'all' | 'offline' | 'online' | 'talk' | 'b2b';
export type SalesDateRange = 'all' | 'today' | 'week' | 'month';

/** 오프라인 판매 목록 */
export function useSales(filters?: {
  search?: string;
  page?: number;
  limit?: number;
  tab?: SalesTab;
  channel?: SalesChannel;
  dateRange?: SalesDateRange;
  dateFrom?: string;
  dateTo?: string;
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
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let query = (supabase as any)
        .from('offline_sales')
        .select('*', { count: 'exact' })
        .order('sale_date', { ascending: false })
        .order('created_at', { ascending: false })
        .range(from, to);

      // 탭 필터 — 전체/오늘/미수금은 취소 건 제외, 취소 탭만 취소 건 표시
      const tab = filters?.tab || 'all';
      if (tab === 'today') {
        const today = new Date().toISOString().slice(0, 10);
        query = query.eq('sale_date', today).is('cancelled_at', null);
      } else if (tab === 'unpaid') {
        query = query.in('payment_status', ['unpaid', 'partial']).is('cancelled_at', null);
      } else if (tab === 'cancelled') {
        query = query.not('cancelled_at', 'is', null);
      } else {
        // all: 취소 건 제외
        query = query.is('cancelled_at', null);
      }

      // 채널 필터
      if (filters?.channel && filters.channel !== 'all') {
        if (filters.channel === 'b2b') {
          query = query.in('customer_type', ['dealer', 'academy']);
        } else {
          query = query.eq('sale_channel', filters.channel);
        }
      }

      // 기간 필터: 커스텀 날짜 범위 우선, 없으면 프리셋
      if (filters?.dateFrom) {
        query = query.gte('sale_date', filters.dateFrom);
      } else if (filters?.dateRange && filters.dateRange !== 'all') {
        const now = new Date();
        let df: string;
        if (filters.dateRange === 'today') {
          df = now.toISOString().slice(0, 10);
        } else if (filters.dateRange === 'week') {
          const d = new Date(now);
          d.setDate(d.getDate() - 7);
          df = d.toISOString().slice(0, 10);
        } else {
          const d = new Date(now);
          d.setMonth(d.getMonth() - 1);
          df = d.toISOString().slice(0, 10);
        }
        query = query.gte('sale_date', df);
      }
      if (filters?.dateTo) {
        query = query.lte('sale_date', filters.dateTo);
      }

      // 검색
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

/** 판매 탭별 건수 */
export function useSalesTabCounts() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sales-tab-counts'],
    staleTime: 30_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const today = new Date().toISOString().slice(0, 10);

      const [allRes, todayRes, unpaidRes, cancelledRes] = await Promise.all([
        db.from('offline_sales').select('*', { count: 'exact', head: true }).is('cancelled_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).eq('sale_date', today).is('cancelled_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).in('payment_status', ['unpaid', 'partial']).is('cancelled_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).not('cancelled_at', 'is', null),
      ]);

      return {
        all: allRes.count || 0,
        today: todayRes.count || 0,
        unpaid: unpaidRes.count || 0,
        cancelled: cancelledRes.count || 0,
      };
    },
  });
}

/** 판매 통계 (주간/월간 매출·건수·미수금) */
export function useSalesStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sales-stats'],
    staleTime: 30_000, // 075: 60s → 30s (대시보드 즉각 반영)
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const now = new Date();
      // 로컬(KST) 기준 ISO date — toISOString은 UTC라 KST 자정 직후 4/30로 잘못 변환되는 버그 회피
      const pad2 = (n: number) => String(n).padStart(2, '0');
      const today = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;

      // 이번주 월요일
      const dayOfWeek = now.getDay();
      const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      const monday = new Date(now);
      monday.setDate(now.getDate() + mondayOffset);
      const weekStart = `${monday.getFullYear()}-${pad2(monday.getMonth() + 1)}-${pad2(monday.getDate())}`;

      // 이번달 1일
      const monthStart = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-01`;

      // 주간 매출
      const { data: weekSales } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount')
        .gte('sale_date', weekStart)
        .lte('sale_date', today)
        .is('cancelled_at', null);

      // 월간 매출
      const { data: monthSales } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount')
        .gte('sale_date', monthStart)
        .lte('sale_date', today)
        .is('cancelled_at', null);

      // 미수금 총액
      const { data: unpaidSales } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount')
        .in('payment_status', ['unpaid', 'partial'])
        .is('cancelled_at', null);

      // 발생주의: 매출 = total_amount - discount_amount (미수금 포함)
      const sum = (arr: Array<{ total_amount: number; paid_amount: number; discount_amount: number }> | null) => {
        if (!arr) return { amount: 0, count: 0 };
        const amount = arr.reduce((s, r) => s + ((r.total_amount || 0) - (r.discount_amount || 0)), 0);
        return { amount, count: arr.length };
      };

      const unpaidTotal = (unpaidSales || []).reduce(
        (s: number, r: { total_amount: number; paid_amount: number; discount_amount: number }) =>
          s + ((r.total_amount - (r.discount_amount || 0)) - (r.paid_amount || 0)),
        0
      );

      // B2B 거래처별 이번달 매출
      const { data: b2bSales } = await db
        .from('offline_sales')
        .select('customer_name, customer_type, total_amount, discount_amount')
        .in('customer_type', ['dealer', 'academy'])
        .gte('sale_date', monthStart)
        .lte('sale_date', today)
        .is('cancelled_at', null);

      const b2bMap: Record<string, { name: string; type: string; amount: number; count: number }> = {};
      for (const s of (b2bSales || [])) {
        const key = s.customer_name || '미지정';
        if (!b2bMap[key]) b2bMap[key] = { name: key, type: s.customer_type || '', amount: 0, count: 0 };
        b2bMap[key].amount += (s.total_amount || 0) - (s.discount_amount || 0);
        b2bMap[key].count++;
      }
      const b2bRanking = Object.values(b2bMap).sort((a, b) => b.amount - a.amount);

      return {
        week: sum(weekSales),
        month: sum(monthSales),
        outstanding: unpaidTotal,
        b2b: b2bRanking,
      };
    },
  });
}

/** 오프라인 판매 단건 조회 (항목 + 시리얼 + 070: 원본 상담 link 포함) */
export function useSale(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sale', id],
    queryFn: async () => {
      const [saleRes, itemsRes, serialsRes] = await Promise.all([
        supabase.from('offline_sales').select('*').eq('id', id).single(),
        supabase.from('offline_sale_items').select('*').eq('sale_id', id),
        supabase.from('product_serials').select('id, serial_number, product_id, sale_item_id').eq('offline_sale_id', id),
      ]);
      if (saleRes.error) throw saleRes.error;
      const sale = saleRes.data as OfflineSale;

      // customer_id가 있으면 최신 연락처 가져오기
      if (sale.customer_id) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data: cust } = await (supabase as any)
          .from('customers')
          .select('phone')
          .eq('id', sale.customer_id)
          .single();
        if (cust?.phone) {
          (sale as Record<string, unknown>).customer_phone = cust.phone;
        }
      }

      // 070: source_consultation_id 있으면 원본 상담 정보 조회 (mirror 모드용)
      let linkedConsultation: { id: string; unique_id: string; name: string; status: string; consultation_type: string } | null = null;
      const sourceId = (sale as { source_consultation_id?: string | null }).source_consultation_id;
      if (sourceId) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data: consult } = await (supabase as any)
          .from('consultations')
          .select('id, unique_id, name, status, consultation_type')
          .eq('id', sourceId)
          .single();
        if (consult) linkedConsultation = consult;
      }

      return {
        sale,
        items: (itemsRes.data || []) as OfflineSaleItem[],
        // sale_item_id 포함 — 상세 페이지에서 정확한 시리얼-상품 매칭에 사용 (2026-05-17 fix)
        serials: (serialsRes.data || []) as Array<{ id: string; serial_number: string; product_id: string | null; sale_item_id: string | null }>,
        linkedConsultation,
      };
    },
    enabled: !!id,
  });
}

/** 제품 목록 (판매 입력용) */
export function useProducts(opts?: { includeInactive?: boolean }) {
  const supabase = createClient();
  const includeInactive = opts?.includeInactive || false;

  return useQuery({
    queryKey: ['products', { includeInactive }],
    staleTime: 60 * 1000, // 075: 5분 → 1분 (시리얼 sold/재고 변경 후 빠른 반영)
    queryFn: async () => {
      let query = supabase
        .from('products')
        .select('*')
        .order('category')
        .order('name');
      if (!includeInactive) query = query.eq('is_active', true);
      const { data, error } = await query;
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
        payment_detail?: Record<string, number>;
        customer_type?: string;
        memo?: string;
        sale_channel?: string;
        contract_id?: string | null;
        source_consultation_id?: string;        // 070: 상담 → 판매 link
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
      /** Phase A — 시리얼 다른 판매에서 이전 명시 동의 (2026-05-18) */
      allow_serial_transfer?: boolean;
    }) => {
      const res = await fetch('/api/sales', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '판매 등록 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`판매 등록 완료: ${data.saleNumber}`);
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('판매 등록 실패: ' + (err instanceof Error ? err.message : String(err)));
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
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('판매 취소 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 반품 처리 */
export function useReturnSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, reason }: { id: string; reason?: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'return', reason }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '반품 처리 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('반품 처리 완료 — 재고/시리얼 복원됨');
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('반품 처리 실패: ' + (err instanceof Error ? err.message : String(err)));
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
      toast.error('결제상태 변경 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
    onSuccess: () => {
      toast.success('결제상태가 변경되었습니다');
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
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
      toast.error('메모 수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
    },
  });
}

/** 판매 정보 수정 (금액/할인/결제방법/날짜) */
export function useEditSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...fields }: {
      id: string;
      total_amount?: number;
      discount_amount?: number;
      payment_method?: string;
      sale_date?: string;
      payment_detail?: Record<string, number>;
    }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'edit_sale', ...fields }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '수정 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('판매 정보가 수정되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 판매 재구성 (제품 추가/삭제 — 시리얼/재고 자동 재조정) */
export function useRebuildSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, items, sale_info, allow_serial_transfer }: {
      id: string;
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        serial_ids?: string[];
        manual_serials?: string[];
      }>;
      sale_info: {
        total_amount: number;
        discount_amount?: number;
        payment_method: string;
        payment_status?: string;
        paid_amount?: number;
        sale_date?: string;
        payment_detail?: Record<string, number>;
        memo?: string;
        sale_channel?: string;
      };
      /** Phase A — 시리얼 다른 판매에서 이전 명시 동의 (2026-05-18) */
      allow_serial_transfer?: boolean;
    }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'rebuild_sale', items, sale_info, allow_serial_transfer }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '수정 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('판매가 수정되었습니다 (시리얼/재고 재조정 완료)');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('판매 수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 판매 송장 생성 (ALPS) */
export function useShipSale() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/sales/${id}/ship`, { method: 'POST' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '송장 생성 실패');
      }
      return res.json();
    },
    onSuccess: (data, id) => {
      toast.success(`송장 생성 완료: ${data.invoiceNumber}`);
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 판매 송장 취소 */
export function useCancelSaleShipment() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/sales/${id}/ship`, { method: 'DELETE' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '송장 취소 실패');
      }
      return res.json();
    },
    onSuccess: (_d, id) => {
      toast.success('송장 취소 완료');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 판매 출고완료 처리 (shipped_at 설정 + 선택적 알림톡) */
export function useMarkSaleShipped() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, send_notification }: { id: string; send_notification: boolean }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'mark_shipped', send_notification }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '출고완료 처리 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('출고완료 처리되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/**
 * 판매 배송완료 처리 — 두 경로 통합 (2026-05-25)
 *   1. mode='delivery': 송장 있는 택배 발송 케이스 — ALPS 추적 실패 fallback (수동 배송완료)
 *   2. mode='pickup':   송장 없는 매장 직접 수령 케이스 — "고객 수령 완료" 처리
 *   두 경로 모두 offline_sales.delivered_at 설정. invoice_number 유무로 칩 표시 분기.
 */
export function useMarkSaleDelivered() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, mode }: { id: string; mode: 'delivery' | 'pickup' }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'mark_delivered', mode }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '배송완료 처리 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id, mode }) => {
      toast.success(mode === 'pickup' ? '고객 수령 완료 처리되었습니다' : '배송완료 처리되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      queryClient.invalidateQueries({ queryKey: ['sales'] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

