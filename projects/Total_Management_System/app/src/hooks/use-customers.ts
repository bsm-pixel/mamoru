'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
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
    staleTime: 30 * 1000, // 30초 캐시
  });
}

/** 고객 신규 등록 (TMS + 이카운트 동시) */
export function useCreateCustomer() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      name: string;
      phone?: string;
      email?: string;
      address?: string;
      postcode?: string;
      addressDetail?: string;
    }) => {
      const res = await fetch('/api/customers', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        customer: CustomerResult;
        ecountSynced: boolean;
        ecountCode: string | null;
      }>;
    },
    onSuccess: (data) => {
      if (data.ecountSynced) {
        toast.success(`고객 등록 완료 (이카운트: ${data.ecountCode})`);
      } else {
        toast.success('고객 등록 완료 (이카운트 미연동)');
      }
      queryClient.invalidateQueries({ queryKey: ['customer-search'] });
    },
    onError: (err) => {
      toast.error('고객 등록 실패: ' + String(err));
    },
  });
}

/** 이카운트 거래처 → TMS 일괄 동기화 */
export function useSyncEcountCustomers() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/customers/sync-ecount', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        total: number;
        existing: number;
        inserted: number;
        message: string;
      }>;
    },
    onSuccess: (data) => {
      toast.success(data.message);
      queryClient.invalidateQueries({ queryKey: ['customer-search'] });
    },
    onError: (err) => {
      toast.error('이카운트 동기화 실패: ' + String(err));
    },
  });
}
