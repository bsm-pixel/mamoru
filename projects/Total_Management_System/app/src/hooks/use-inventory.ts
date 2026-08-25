'use client';

import { useQuery } from '@tanstack/react-query';

export interface InventoryItem {
  id: string;
  name: string;
  sku: string;
  category: string;
  price: number;
  price_dealer: number;
  price_purchase: number;
  price_groups: Record<string, { price?: number | null; display_name?: string | null }> | null;
  barcode: string | null;
  stock_quantity: number;
  raw_stock: number;
  pending_quantity: number;
  zone_raw: number;
  zone_ready: number;
  zone_display: number;
  /** 반품창고(status=returned) — 판매가능 현재고와 분리, 무결식 합계에 미포함 */
  zone_return?: number;
  /** 112: 정위치 (표시 전용, 미지정이면 null) */
  location_id: string | null;
  location_code: string | null;
  location_label: string | null;
}

export interface InventorySummary {
  total_products: number;
  total_stock: number;
  total_pending: number;
  low_stock_count: number;
  total_value: number;
}

/** 재고 현황 조회 */
export function useInventory(filters?: {
  category?: string;
  search?: string;
  lowStock?: boolean;
  threshold?: number;
}) {
  const params = new URLSearchParams();
  if (filters?.category) params.set('category', filters.category);
  if (filters?.search) params.set('search', filters.search);
  if (filters?.lowStock) params.set('low_stock', 'true');
  if (filters?.threshold) params.set('threshold', String(filters.threshold));

  return useQuery<{ items: InventoryItem[]; summary: InventorySummary }>({
    queryKey: ['inventory', filters],
    queryFn: async () => {
      const res = await fetch(`/api/inventory?${params.toString()}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
  });
}
