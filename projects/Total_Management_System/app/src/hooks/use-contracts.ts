'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Contract, ContractItem } from '@/lib/supabase/types';
import toast from 'react-hot-toast';

/** 계약서 탭 타입 */
export type ContractTab = 'all' | 'new' | 'converted' | 'cancelled';

/** 계약서 목록 */
export function useContracts(filters?: {
  search?: string;
  status?: string;
  tab?: ContractTab;
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
    queryKey: ['contracts', filters],
    staleTime: 30_000,
    queryFn: async () => {
      let query = supabase
        .from('contracts')
        .select('*', { count: 'exact' })
        .order('created_at', { ascending: false })
        .range(from, to);

      // 탭 필터
      const tab = filters?.tab || 'all';
      if (tab === 'new') {
        // 신규계약: signed/sent + 판매 전환 안 됨
        query = query.in('status', ['draft', 'signed', 'sent']).is('offline_sale_id', null);
      } else if (tab === 'converted') {
        // 전환완료: offline_sale_id 있음
        query = query.not('offline_sale_id', 'is', null);
      } else if (tab === 'cancelled') {
        query = query.eq('status', 'cancelled');
      }

      // 레거시 status 필터 (다른 곳에서 사용 시)
      if (filters?.status && filters.status !== 'all' && !filters?.tab) {
        query = query.eq('status', filters.status);
      }

      if (filters?.search) {
        query = query.or(
          `customer_name.ilike.%${filters.search}%,customer_phone.ilike.%${filters.search}%,contract_number.ilike.%${filters.search}%`
        );
      }
      if (filters?.dateRange && filters.dateRange !== 'all') {
        const now = new Date();
        let dateFrom: string;
        if (filters.dateRange === 'today') dateFrom = now.toISOString().slice(0, 10);
        else if (filters.dateRange === 'week') { const d = new Date(now); d.setDate(d.getDate() - 7); dateFrom = d.toISOString().slice(0, 10); }
        else { const d = new Date(now); d.setMonth(d.getMonth() - 1); dateFrom = d.toISOString().slice(0, 10); }
        query = query.gte('created_at', dateFrom);
      }

      const { data, count, error } = await query;
      if (error) throw error;
      return { contracts: (data || []) as Contract[], total: count || 0 };
    },
  });
}

/** 계약서 탭별 건수 */
export function useContractTabCounts() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['contract-tab-counts'],
    staleTime: 30_000,
    queryFn: async () => {
      const [allRes, newRes, convertedRes, cancelledRes] = await Promise.all([
        supabase.from('contracts').select('*', { count: 'exact', head: true }),
        supabase.from('contracts').select('*', { count: 'exact', head: true })
          .in('status', ['draft', 'signed', 'sent']).is('offline_sale_id', null),
        supabase.from('contracts').select('*', { count: 'exact', head: true })
          .not('offline_sale_id', 'is', null),
        supabase.from('contracts').select('*', { count: 'exact', head: true })
          .eq('status', 'cancelled'),
      ]);

      return {
        all: allRes.count || 0,
        new: newRes.count || 0,
        converted: convertedRes.count || 0,
        cancelled: cancelledRes.count || 0,
      };
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
        // 필기 + 상담 연결
        consultation_id?: string;
        handwriting_name?: string;
        handwriting_phone?: string;
        handwriting_address?: string;
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
