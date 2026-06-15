'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import type { EventSubmission } from '@/lib/event/types';

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
