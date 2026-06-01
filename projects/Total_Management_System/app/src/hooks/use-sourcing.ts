'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';

export type SourcingInspectionStatus = 'pending' | 'matched' | 'selected' | 'rejected';

export interface SourcingItem {
  id: string;
  po_id: string;
  sticker_no: string;
  supplier_name: string | null;
  supplier_url: string | null;
  vendor_url: string | null;
  product_name: string;
  features_memo: string | null;
  unit_price: number;
  moq: number | null;
  inbound_photos: string[];
  inbound_memo: string | null;
  inspection_status: SourcingInspectionStatus;
  selected_at: string | null;
  sort_order: number;
  linked_product_id: string | null;
  linked_product?: { id: string; sku: string; name: string; category: string } | null;
}

export interface SourcingPo {
  id: string;
  po_number: string;
  supplier_name: string | null;
  supplier_url: string | null;
  order_date: string;
  exchange_rate: number;
  status: string;
  memo: string | null;
  created_at: string;
}

export interface SourcingPoListItem extends SourcingPo {
  counts: { total: number; pending: number; matched: number; selected: number; rejected: number };
}

/** 소싱 발주 목록 */
export function useSourcingList(status?: string) {
  return useQuery({
    queryKey: ['sourcing-list', status ?? ''],
    queryFn: async () => {
      const params = new URLSearchParams();
      if (status) params.set('status', status);
      const res = await fetch(`/api/sourcing?${params}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ orders: SourcingPoListItem[] }>;
    },
  });
}

/** 소싱 발주 상세 (+ 품목) */
export function useSourcingDetail(id: string) {
  return useQuery({
    queryKey: ['sourcing-detail', id],
    queryFn: async () => {
      const res = await fetch(`/api/sourcing/${id}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ po: SourcingPo; items: SourcingItem[] }>;
    },
    enabled: !!id,
  });
}

/** 소싱 발주 생성 (빈 발주 또는 품목 포함) */
export function useCreateSourcing() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (data: {
      supplier_name?: string;
      supplier_url?: string;
      order_date?: string;
      exchange_rate?: number;
      memo?: string;
      items?: Array<{ vendor_url?: string; product_name?: string; features_memo?: string; unit_price?: number; moq?: number | null }>;
    }) => {
      const res = await fetch('/api/sourcing', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ id: string; po_number: string }>;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-list'] }),
    onError: (e) => toast.error('소싱 생성 실패: ' + String(e)),
  });
}

/** 발주 헤더 수정 */
export function useUpdateSourcingPo() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, ...patch }: { id: string } & Partial<SourcingPo>) => {
      const res = await fetch(`/api/sourcing/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(patch),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (_d, v) => {
      qc.invalidateQueries({ queryKey: ['sourcing-detail', v.id] });
      qc.invalidateQueries({ queryKey: ['sourcing-list'] });
    },
    onError: (e) => toast.error('수정 실패: ' + String(e)),
  });
}

/** 발주 삭제 */
export function useDeleteSourcingPo() {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/sourcing/${id}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('삭제 완료');
      qc.invalidateQueries({ queryKey: ['sourcing-list'] });
    },
    onError: (e) => toast.error('삭제 실패: ' + String(e)),
  });
}

/** 품목 추가 */
export function useAddSourcingItem(poId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (data?: {
      supplier_name?: string | null;
      supplier_url?: string | null;
      vendor_url?: string | null;
      product_name?: string;
      features_memo?: string | null;
      unit_price?: number;
      moq?: number | null;
    }) => {
      const res = await fetch(`/api/sourcing/${poId}/items`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data ?? {}),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ item: SourcingItem }>;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-detail', poId] }),
    onError: (e) => toast.error('품목 추가 실패: ' + String(e)),
  });
}

/** 품목 수정 (입력값 / 선별 상태 / 사진·메모) */
export function useUpdateSourcingItem(poId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ itemId, ...patch }: { itemId: string } & Partial<Omit<SourcingItem, 'id' | 'po_id' | 'sticker_no'>>) => {
      const res = await fetch(`/api/sourcing/items/${itemId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(patch),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ item: SourcingItem }>;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-detail', poId] }),
    onError: (e) => toast.error('품목 수정 실패: ' + String(e)),
  });
}

/** 품목 ↔ 제품/부자재 연결 (linked_product_id set/null) */
export function useLinkSourcingItem(poId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ itemId, linked_product_id }: { itemId: string; linked_product_id: string | null }) => {
      const res = await fetch(`/api/sourcing/items/${itemId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ linked_product_id }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-detail', poId] }),
    onError: (e) => toast.error('연결 실패: ' + String(e)),
  });
}

// ── 품목 이미지 (STEP 1 주문 시 붙여넣기/업로드 — 서버 영구 저장) ──
/** 품목 이미지 업로드 (상세 페이지용 — 1688 이미지 붙여넣기/업로드) */
export function useUploadItemImage(poId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ itemId, file }: { itemId: string; file: File }) => {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch(`/api/sourcing/items/${itemId}/photos`, { method: 'POST', body: fd });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ url: string; inbound_photos: string[] }>;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-detail', poId] }),
    onError: (e) => toast.error('이미지 업로드 실패: ' + String(e)),
  });
}

/** 품목 이미지 삭제 (상세 페이지용) */
export function useDeleteItemImage(poId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async ({ itemId, url }: { itemId: string; url: string }) => {
      const res = await fetch(`/api/sourcing/items/${itemId}/photos?url=${encodeURIComponent(url)}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-detail', poId] }),
    onError: (e) => toast.error('이미지 삭제 실패: ' + String(e)),
  });
}

// ── 모바일 입고매칭 (단건 품목) ──────────────────────────

/** 단건 품목 조회 (QR 스캔 직진입) */
export function useSourcingItem(itemId: string) {
  return useQuery({
    queryKey: ['sourcing-item', itemId],
    queryFn: async () => {
      const res = await fetch(`/api/sourcing/items/${itemId}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ item: SourcingItem }>;
    },
    enabled: !!itemId,
  });
}

/** 입고 품목 수정 (메모/매칭상태) — 품목 단건 캐시 무효화 */
export function useInboundItemUpdate(itemId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (patch: { inbound_memo?: string; inspection_status?: SourcingInspectionStatus }) => {
      const res = await fetch(`/api/sourcing/items/${itemId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(patch),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ item: SourcingItem }>;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-item', itemId] }),
    onError: (e) => toast.error('저장 실패: ' + String(e)),
  });
}

/** 입고 사진 업로드 (카메라) */
export function useUploadInboundPhoto(itemId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (file: File) => {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch(`/api/sourcing/items/${itemId}/photos`, { method: 'POST', body: fd });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ url: string; inbound_photos: string[] }>;
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-item', itemId] }),
    onError: (e) => toast.error('사진 업로드 실패: ' + String(e)),
  });
}

/** 입고 사진 삭제 */
export function useDeleteInboundPhoto(itemId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (url: string) => {
      const res = await fetch(`/api/sourcing/items/${itemId}/photos?url=${encodeURIComponent(url)}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-item', itemId] }),
    onError: (e) => toast.error('사진 삭제 실패: ' + String(e)),
  });
}

/** 품목 삭제 */
export function useDeleteSourcingItem(poId: string) {
  const qc = useQueryClient();
  return useMutation({
    mutationFn: async (itemId: string) => {
      const res = await fetch(`/api/sourcing/items/${itemId}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => qc.invalidateQueries({ queryKey: ['sourcing-detail', poId] }),
    onError: (e) => toast.error('품목 삭제 실패: ' + String(e)),
  });
}
