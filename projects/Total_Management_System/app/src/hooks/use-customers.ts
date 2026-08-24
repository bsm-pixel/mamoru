'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { Customer } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

export interface CustomerResult {
  id: string;
  name: string;
  phone: string | null;
  email: string | null;
  address_road: string | null;
  address_detail: string | null;
  postcode: string | null;
  ecount_customer_code: string | null;
  source: string;
  memo: string | null;
  company_name: string | null;
  activity_name: string | null;          // 102: 활동명(매장 사용 이름)
  position: string | null;               // 102: 직급
  customer_type: string | null;
  default_repair_price: number | null;  // 079: 거래처별 복원수리 기본 단가(자루당)
}

/** 고객 자동완성 검색 (2글자 이상) */
export function useCustomerSearch(query: string) {
  return useQuery({
    queryKey: ['customer-search', query],
    queryFn: async () => {
      const res = await fetch(`/api/customers/search?q=${encodeURIComponent(query)}`);
      if (!res.ok) throw new Error(await res.text());
      const data = await res.json();
      return (data.customers || []) as CustomerResult[];
    },
    enabled: query.length >= 2,
    staleTime: 30 * 1000,
  });
}

/** 고객 목록 (검색 + 유형 필터 + 페이징) */
export function useCustomers(filters?: {
  search?: string;
  type?: string;
  page?: number;
  limit?: number;
}) {
  return useQuery({
    queryKey: ['customers', filters],
    staleTime: 30_000,
    queryFn: async () => {
      const params = new URLSearchParams();
      if (filters?.search) params.set('search', filters.search);
      if (filters?.type) params.set('type', filters.type);
      if (filters?.page) params.set('page', String(filters.page));
      if (filters?.limit) params.set('limit', String(filters.limit));

      const res = await fetch(`/api/customers?${params}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ customers: Customer[]; total: number }>;
    },
  });
}

/** 고객 상세 (판매내역 + 계약서 + 상담 포함) */
export function useCustomer(id: string) {
  return useQuery({
    queryKey: ['customer', id],
    queryFn: async () => {
      const res = await fetch(`/api/customers/${id}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        customer: Customer;
        sales: Array<{
          id: string;
          sale_number: string;
          sale_date: string;
          total_amount: number;
          paid_amount: number;
          payment_method: string;
          payment_status: string;
        }>;
        contracts: Array<{
          id: string;
          contract_number: string;
          final_amount: number;
          status: string;
          created_at: string;
        }>;
        consultations: Array<{
          id: string;
          consultation_type: string;
          visit_date: string | null;
          status: string;
          created_at: string;
        }>;
        repairs: Array<{
          id: string;
          repair_number: string;
          status: string;
          item_description: string | null;
          total_cost: number | null;
          created_at: string;
        }>;
        orders?: Array<{
          id: string;
          imweb_order_no: string;
          ordered_at: string;
          paid_amount: number | null;
          status: string;
        }>;
        summary: {
          totalSales: number;
          totalSalesAmount: number;
        };
      }>;
    },
    enabled: !!id,
  });
}

/** 고객 신규 등록 */
export function useCreateCustomer() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      name: string;
      phone?: string;
      email?: string;
      postcode?: string;
      address_road?: string;
      address_detail?: string;
      customer_type?: string;
      company_name?: string;
      memo?: string;
      tags?: string[];
    }) => {
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) {
        const j = await res.json().catch(() => null);
        // 409: 같은 전화 중복 — 기존 고객명 안내
        if (res.status === 409 && j?.existing) throw new Error(`이미 등록된 번호입니다 (${j.existing.name})`);
        throw new Error(j?.error ? String(j.error) : '등록 실패');
      }
      return res.json() as Promise<{ customer: CustomerResult }>;
    },
    onSuccess: () => {
      toast.success('고객 등록 완료');
      queryClient.invalidateQueries({ queryKey: ['customer-search'] });
      queryClient.invalidateQueries({ queryKey: ['customers'] });
    },
    onError: (err) => {
      toast.error('고객 등록 실패: ' + String(err));
    },
  });
}

/** 고객 병합 — 흡수 대상들을 주 고객으로 통합 (거래·미수금 이관) */
export function useMergeCustomers() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ primaryId, victimIds }: { primaryId: string; victimIds: string[] }) => {
      const res = await fetch('/api/customers/merge', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ primaryId, victimIds }),
      });
      const json = await res.json();
      if (!res.ok) throw new Error(json.error || '병합 실패');
      return json as { ok: true; result: { primary: string; victims: number } };
    },
    onSuccess: () => {
      // 고객·거래·집계 전반 갱신
      queryClient.invalidateQueries({ queryKey: ['customers'] });
      queryClient.invalidateQueries({ queryKey: ['customer'] });
      queryClient.invalidateQueries({ queryKey: ['customer-search'] });
      queryClient.invalidateQueries({ queryKey: ['sales-stats'] });
      queryClient.invalidateQueries({ queryKey: ['delivery-stats'] });
    },
  });
}

/** 고객 정보 수정 */
export function useUpdateCustomer() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...data }: {
      id: string;
      name?: string;
      phone?: string;
      email?: string;
      postcode?: string;
      address_road?: string;
      address_detail?: string;
      customer_type?: string;
      company_name?: string;
      activity_name?: string | null;
      position?: string | null;
      memo?: string;
      outstanding_balance?: number;
      default_repair_price?: number | null;
    }) => {
      const res = await fetch(`/api/customers/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ customer: Customer }>;
    },
    onSuccess: (_, vars) => {
      toast.success('고객 정보 수정 완료');
      queryClient.invalidateQueries({ queryKey: ['customer', vars.id] });
      queryClient.invalidateQueries({ queryKey: ['customers'] });
      queryClient.invalidateQueries({ queryKey: ['customer-search'] });
    },
    onError: (err) => {
      toast.error('수정 실패: ' + String(err));
    },
  });
}
