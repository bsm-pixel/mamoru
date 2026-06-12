/**
 * 바코드 스캔값 해석 — SKU(품목) vs 시리얼(MR형식) 자동 분기.
 * 시리얼 = 경량 API 조회, SKU = 로드된 제품에서 sku/barcode 매칭(빠름).
 */

import { isMSerial } from '@/lib/serial/format';
import type { Product } from '@/lib/supabase/types';

export interface ScanSerial {
  id: string;
  serial_number: string;
  product_id: string | null;
  status: string;
  offline_sale_id: string | null;
  sale_item_id: string | null;
  sold_to_name: string | null;
}

export type ScanResult =
  | { type: 'product'; product: Product }
  | { type: 'serial'; serial: ScanSerial; product: Product | null }
  | { type: 'serial-notfound'; code: string }
  | { type: 'notfound'; code: string };

export async function resolveScan(code: string, products: Product[]): Promise<ScanResult> {
  const v = code.trim();
  if (!v) return { type: 'notfound', code: v };

  // 시리얼(MR형식) → 경량 조회
  if (isMSerial(v)) {
    try {
      const r = await fetch(`/api/serials/by-code?code=${encodeURIComponent(v)}`);
      const j = await r.json();
      if (!j?.found) return { type: 'serial-notfound', code: v };
      const product = products.find((p) => p.id === j.serial.product_id) || null;
      return { type: 'serial', serial: j.serial as ScanSerial, product };
    } catch {
      return { type: 'serial-notfound', code: v };
    }
  }

  // SKU/바코드 → 로드된 제품에서 매칭
  const lower = v.toLowerCase();
  const product = products.find(
    (p) => (p.sku || '').toLowerCase() === lower || ((p.barcode || '') as string).toLowerCase() === lower,
  );
  return product ? { type: 'product', product } : { type: 'notfound', code: v };
}
