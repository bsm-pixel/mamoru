'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';

/** 112: 창고 로케이션(정위치) — 2026-07-18 */

export interface LocationProduct {
  id: string;
  name: string;
  sku: string;
  category: string;
  stock_quantity: number;
  location_id: string | null;
}

export interface LocationWithProducts {
  id: string;
  code: string;
  label: string | null;
  rack_no: number;
  level_no: number;
  bin_no: number | null;
  zone_type: string;
  sort_order: number;
  is_active: boolean;
  memo: string | null;
  /** 이 칸에 배정된 제품 수 */
  product_count: number;
  /** 이 칸 제품들의 재고 합 (읽기 전용 집계) */
  stock_total: number;
  products: LocationProduct[];
}

/** 113: 렉 정보 (렉 전체 열 수) */
export interface RackInfo {
  id: string;
  rack_no: number;
  label: string | null;
  columns: number;
  sort_order: number;
  memo: string | null;
}

interface LocationsResponse {
  racks: RackInfo[];
  locations: LocationWithProducts[];
  unassigned: LocationProduct[];
  unassigned_count: number;
  total_locations: number;
}

/** 로케이션 목록 + 칸별 배정 제품 */
export function useLocations() {
  return useQuery({
    queryKey: ['warehouse-locations'],
    staleTime: 30_000,
    queryFn: async (): Promise<LocationsResponse> => {
      const res = await fetch('/api/warehouse/locations');
      if (!res.ok) throw new Error('로케이션 조회 실패');
      return res.json();
    },
  });
}

/** 공용 — POST action 호출 + 캐시 갱신 */
function useWarehouseAction(successMsg: (d: Record<string, unknown>) => string) {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (payload: Record<string, unknown>) => {
      const res = await fetch('/api/warehouse/locations', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      const d = await res.json().catch(() => ({ error: res.statusText }));
      if (!res.ok) throw new Error(d.error || '처리 실패');
      return d;
    },
    onSuccess: (d) => {
      toast.success(successMsg(d));
      queryClient.invalidateQueries({ queryKey: ['warehouse-locations'] });
      queryClient.invalidateQueries({ queryKey: ['inventory'] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : String(e)),
  });
}

/** 렉 생성 — 단별 칸 수 배열 (예: [2,6,6,0,0]) */
export function useCreateRack() {
  const m = useWarehouseAction((d) => `렉 생성 완료 — ${d.created}칸`);
  return {
    ...m,
    mutate: (v: { rack_no: number; label?: string | null; level_bins: number[]; zone_type?: string }, opts?: Parameters<typeof m.mutate>[1]) =>
      m.mutate({ action: 'create_rack', ...v }, opts),
  };
}

/** 특정 단에 칸 1개 추가 */
export function useAddBin() {
  const m = useWarehouseAction(() => '칸을 추가했습니다');
  return { ...m, mutate: (v: { rack_no: number; level_no: number }) => m.mutate({ action: 'add_bin', ...v }) };
}

/** 렉에 단 1개 추가 (기본은 칸 없이 선반) */
export function useAddLevel() {
  const m = useWarehouseAction((d) => `${d.level_no}단을 추가했습니다`);
  return { ...m, mutate: (v: { rack_no: number; bins?: number }) => m.mutate({ action: 'add_level', ...v }) };
}

/** 렉/칸 삭제 — 배정돼 있던 제품은 '미지정'으로 풀린다 */
export function useDeleteLocation() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, rack_no }: { id?: string; rack_no?: number }) => {
      const qs = id ? `id=${encodeURIComponent(id)}` : `rack_no=${rack_no}`;
      const res = await fetch(`/api/warehouse/locations?${qs}`, { method: 'DELETE' });
      const d = await res.json().catch(() => ({ error: res.statusText }));
      if (!res.ok) throw new Error(d.error || '삭제 실패');
      return d;
    },
    onSuccess: () => {
      toast.success('삭제되었습니다 (배정됐던 제품은 미지정으로 바뀝니다)');
      queryClient.invalidateQueries({ queryKey: ['warehouse-locations'] });
      queryClient.invalidateQueries({ queryKey: ['inventory'] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : String(e)),
  });
}

/** 제품 → 위치 배정 (단건·일괄 공용). location_id=null 이면 해제 */
export function useAssignLocation() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ product_ids, location_id }: { product_ids: string[]; location_id: string | null }) => {
      const res = await fetch('/api/warehouse/assign', {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ product_ids, location_id }),
      });
      const d = await res.json().catch(() => ({ error: res.statusText }));
      if (!res.ok) throw new Error(d.error || '위치 배정 실패');
      return d;
    },
    onSuccess: (d, { location_id }) => {
      toast.success(location_id ? `위치 지정 완료 (${d.updated}건)` : `위치 해제 완료 (${d.updated}건)`);
      queryClient.invalidateQueries({ queryKey: ['warehouse-locations'] });
      queryClient.invalidateQueries({ queryKey: ['inventory'] });
    },
    onError: (e) => toast.error(e instanceof Error ? e.message : String(e)),
  });
}
