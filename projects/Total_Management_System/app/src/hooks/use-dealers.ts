'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { Dealer } from '@/lib/supabase/types';

/** 활성 딜러 목록 조회 */
export function useDealers() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['dealers'],
    queryFn: async () => {
      const { data, error } = await supabase
        .from('dealers')
        .select('*')
        .eq('status', 'active')
        .order('name');

      if (error) throw error;
      return (data || []) as Dealer[];
    },
  });
}
