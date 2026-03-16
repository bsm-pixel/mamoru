'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Contract, ContractItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 계약서 목록 */
export function useContracts(filters?: {
  search?: string;
  status?: string;
  page?: number;
  limit?: number;
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 20;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['contracts', filters],
    staleTime: 30_000,
    queryFn: async () => {
      let query = supabase
        .from('contracts')
        .select('*', { count: 'exact' })
        .order('created_at', { ascending: false })
        .range(from, to);

      if (filters?.status && filters.status !== 'all') {
        query = query.eq('status', filters.status);
      }
      if (filters?.search) {
        query = query.or(
          `customer_name.ilike.%${filters.search}%,customer_phone.ilike.%${filters.search}%,contract_number.ilike.%${filters.search}%`
        );
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { contracts: (data || []) as Contract[], total: count || 0 };
    },
  });
}

/** 계약서 단건 조회 */
export function useContract(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['contract', id],
    queryFn: async () => {
      const [contractRes, itemsRes] = await Promise.all([
        supabase.from('contracts').select('*').eq('id', id).single(),
        supabase.from('contract_items').select('*').eq('contract_id', id),
      ]);
      if (contractRes.error) throw contractRes.error;
      return {
        contract: contractRes.data as Contract,
        items: (itemsRes.data || []) as ContractItem[],
      };
    },
    enabled: !!id,
  });
}

/** 계약서 생성 */
export function useCreateContract() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      contract: {
        customer_id?: string;
        customer_name: string;
        customer_phone?: string;
        customer_email?: string;
        customer_address?: string;
        total_amount: number;
        discount_amount?: number;
        final_amount: number;
        payment_method: string;
        installment_months?: number;
        signature_data?: string;
        memo?: string;
        // 전자문서 확장 필드
        delivery_method?: string;
        unavailable_days?: string;
        deposit_amount?: number;
        balance_amount?: number;
        seller_signature?: string;
        customer_title?: string;
        shop_name?: string;
        shop_address?: string;
      };
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        option_text?: string;
      }>;
    }) => {
      const res = await fetch('/api/contracts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`계약서 생성: ${data.contractNumber}`);
      queryClient.invalidateQueries({ queryKey: ['contracts'] });
    },
    onError: (err) => {
      toast.error('계약서 생성 실패: ' + String(err));
    },
  });
}

/** 계약서 알림톡 발송 */
export function useSendContractNotification() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (contractId: string) => {
      const res = await fetch('/api/contracts/notify', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contractId }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('알림톡 발송 완료');
      queryClient.invalidateQueries({ queryKey: ['contracts'] });
      queryClient.invalidateQueries({ queryKey: ['contract'] });
    },
    onError: (err) => {
      toast.error('알림톡 발송 실패: ' + String(err));
    },
  });
}
