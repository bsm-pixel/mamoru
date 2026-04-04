import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';

/* ────────────────────────────────────────────
 * 전체 설정 조회
 * ──────────────────────────────────────────── */
export function useSettings() {
  return useQuery<Record<string, unknown>>({
    queryKey: ['settings'],
    staleTime: 60_000,
    queryFn: async () => {
      const res = await fetch('/api/settings');
      if (!res.ok) throw new Error('설정 조회 실패');
      return res.json();
    },
  });
}

/* ────────────────────────────────────────────
 * 개별 키 조회 + fallback
 * 사용: const goal = useSetting<number>('dashboard.monthly_goal', 0);
 * ──────────────────────────────────────────── */
export function useSetting<T>(key: string, defaultValue: T): T {
  const { data } = useSettings();
  if (!data || data[key] === undefined || data[key] === null) return defaultValue;
  const raw = data[key];
  if (typeof raw === 'string') {
    try {
      return JSON.parse(raw) as T;
    } catch {
      return raw as unknown as T;
    }
  }
  return raw as T;
}

/* ────────────────────────────────────────────
 * 설정 업데이트 뮤테이션
 * 사용: const update = useUpdateSettings();
 *       update.mutate([{ key: 'dashboard.monthly_goal', value: 1500 }]);
 * ──────────────────────────────────────────── */
export function useUpdateSettings() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (items: { key: string; value: unknown }[]) => {
      const res = await fetch('/api/settings', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: '저장 실패' }));
        throw new Error(typeof err.error === 'string' ? err.error : '저장 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('설정 저장 완료');
      queryClient.invalidateQueries({ queryKey: ['settings'] });
    },
    onError: (err: Error) => {
      toast.error('설정 저장 실패: ' + err.message);
    },
  });
}

/* ────────────────────────────────────────────
 * 서버사이드 설정 조회 헬퍼
 * API route에서 사용:
 *   const price = await getServerSetting(supabase, 'repair.price_mamoru', 10000);
 * ──────────────────────────────────────────── */
export async function getServerSetting<T>(
  supabase: any, // eslint-disable-line @typescript-eslint/no-explicit-any
  key: string,
  defaultValue: T,
): Promise<T> {
  try {
    const { data } = await supabase
      .from('system_settings')
      .select('value')
      .eq('key', key)
      .single();
    if (!data?.value) return defaultValue;
    const raw = data.value;
    if (typeof raw === 'string') {
      try {
        return JSON.parse(raw) as T;
      } catch {
        return raw as unknown as T;
      }
    }
    return raw as T;
  } catch {
    return defaultValue;
  }
}