'use client';

import { useQuery } from '@tanstack/react-query';

interface SerialLookupResult {
  serial: {
    id: string;
    serial_number: string;
    barcode: string | null;
    status: string;
    warehouse_zone: string;
    sold_via: string | null;
    sold_at: string | null;
    sold_to_name: string | null;
    sold_to_phone: string | null;
    lot_number: string | null;
    manufactured_at: string | null;
    memo: string | null;
    product_id: string;
    offline_sale_id: string | null;
  } | null;
  product: {
    id: string;
    name: string;
    sku: string;
    category: string;
    image_url: string | null;
    price: number;
  } | null;
  sale: {
    id: string;
    sale_number: string;
    sale_date: string;
    sale_channel: string;
    payment_method: string;
    payment_status: string;
    customer_name: string;
    customer_phone: string | null;
    total_amount: number;
  } | null;
  repairs: Array<{
    id: string;
    status: string;
    created_at: string;
  }>;
}

export function useSerialLookup(query: string) {
  return useQuery<SerialLookupResult>({
    queryKey: ['serial-lookup', query],
    queryFn: async () => {
      const res = await fetch(`/api/serials/lookup?q=${encodeURIComponent(query)}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    enabled: query.length >= 2,
    staleTime: 5 * 60_000,
  });
}

/** 시리얼 이동 이력 — DB 트리거가 자동 기록한 append-only ledger */
export interface SerialAuditLog {
  id: string;
  serial_id: string;
  serial_number: string | null;
  action: 'INSERT' | 'UPDATE' | 'DELETE';
  old_status: string | null;
  new_status: string | null;
  old_warehouse_zone: string | null;
  new_warehouse_zone: string | null;
  old_sale_item_id: string | null;
  new_sale_item_id: string | null;
  old_offline_sale_id: string | null;
  new_offline_sale_id: string | null;
  old_contract_id: string | null;
  new_contract_id: string | null;
  old_product_id: string | null;
  new_product_id: string | null;
  changed_by: string | null;
  changed_by_name: string;
  changed_at: string;
}

export function useSerialAudit(serialId: string | null | undefined) {
  return useQuery<{ logs: SerialAuditLog[] }>({
    queryKey: ['serial-audit', serialId],
    queryFn: async () => {
      const res = await fetch(`/api/serials/audit?serial_id=${serialId}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    enabled: !!serialId,
    staleTime: 30_000,
  });
}
