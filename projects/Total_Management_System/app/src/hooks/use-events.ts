'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { EventSubmission, EventCampaign } from '@/lib/event/types';

export function useCampaigns() {
  return useQuery({
    queryKey: ['campaigns'],
    queryFn: async () => {
      const res = await fetch('/api/campaigns');
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      return (json.campaigns || []) as EventCampaign[];
    },
    staleTime: 60_000,
  });
}

import type { DiscountRule } from '@/lib/event/types';

export function useCreateCampaign() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (body: { name: string; type?: string; discount_rules?: DiscountRule[] }) => {
      const res = await fetch('/api/campaigns', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['campaigns'] }),
  });
}

export function useUpdateCampaign() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, ...body }: { id: string; name?: string; type?: string; status?: string; discount_rules?: DiscountRule[] }) => {
      const res = await fetch(`/api/campaigns/${id}`, {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['campaigns'] }),
  });
}

export function useEvents(status: string = 'all') {
  return useQuery({
    queryKey: ['events', status],
    queryFn: async () => {
      const res = await fetch(`/api/events?status=${status}`);
      if (!res.ok) throw new Error(await res.text());
      const json = await res.json();
      return (json.events || []) as EventSubmission[];
    },
    staleTime: 30_000,
  });
}

interface PatchInput {
  id: string;
  action: 'payment_notice' | 'confirm_payment' | 'cancel' | 'update';
  total_amount?: number;
  memo?: string;
  reason?: string;
  send_notification?: boolean;
}

export function useEventPatch() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, ...body }: PatchInput) => {
      const res = await fetch(`/api/events/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      qc.invalidateQueries({ queryKey: ['events'] });
      qc.invalidateQueries({ queryKey: ['sales'] });
    },
  });
}

/** 접수 기록 완전 삭제 (오등록·테스트 정리용). 판매 건은 별도(판매관리). */
export function useEventDelete() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ id }: { id: string }) => {
      const res = await fetch(`/api/events/${id}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['events'] }),
  });
}
