/**
 * 아임웹 상품 동기화 (v2 API)
 * 아임웹 → TMS products 테이블 upsert
 * 매입가/거래처는 동기화하지 않음 (TMS 자체 관리)
 */

import { getImwebProducts, type ImwebV2Product } from './client';
import { createServiceClient } from '@/lib/supabase/server';

interface SyncResult {
  success: boolean;
  synced: number;
  created: number;
  updated: number;
  errors: string[];
}

/** 아임웹 이미지 URL 생성 */
function getImageUrl(product: ImwebV2Product): string | null {
  if (!product.image_url || !product.images?.[0]) return null;
  const key = product.images[0];
  const path = product.image_url[key];
  if (!path) return null;
  return `https://cdn.imweb.me/upload/${path}`;
}

/** 아임웹 전체 상품을 TMS에 동기화 */
export async function syncProducts(): Promise<SyncResult> {
  const supabase = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = supabase as any;
  const errors: string[] = [];
  let synced = 0;
  let created = 0;
  let updated = 0;

  // sync_log 시작
  const { data: logEntry } = await db.from('sync_log').insert({
    sync_type: 'imweb_products',
    status: 'running',
    started_at: new Date().toISOString(),
  }).select().single();

  try {
    let page = 1;
    let hasMore = true;

    while (hasMore) {
      const res = await getImwebProducts(page, 50);

      if (res.code !== 200 || !res.data?.list) {
        errors.push(`페이지 ${page} 조회 실패: code ${res.code}`);
        break;
      }

      const products = res.data.list;

      for (const p of products) {
        try {
          // 기존 TMS 제품 조회 (imweb_product_no 기준)
          const { data: existing } = await db
            .from('products')
            .select('id, price_purchase, price_dealer, supplier_id, category')
            .eq('imweb_product_no', String(p.no))
            .single();

          // 재고 미사용: stock_use=false → stock_quantity=-1 (TMS 규약)
          const stockQty = p.stock?.stock_use ? (p.stock.stock_no_option || 0) : -1;

          const productData = {
            name: p.name,
            imweb_product_no: String(p.no),
            price: p.price || 0,
            stock_quantity: stockQty,
            image_url: getImageUrl(p),
            sku: p.custom_prod_code || `IW-${p.no}`,
            is_active: p.prod_status === 'sale',
            updated_at: new Date().toISOString(),
          };

          if (existing) {
            // 업데이트 — 매입가/거래처/카테고리는 TMS 값 유지
            await db.from('products').update(productData).eq('id', existing.id);
            updated++;
          } else {
            // 신규 생성
            await db.from('products').insert({
              ...productData,
              category: 'BL', // 기본값, TMS에서 수동 변경
              price_dealer: 0,
              price_purchase: 0,
              created_at: new Date().toISOString(),
            });
            created++;
          }
          synced++;
        } catch (err) {
          errors.push(`상품 ${p.no}(${p.name}): ${String(err)}`);
        }
      }

      // 페이지네이션
      hasMore = page < res.data.pagenation.total_page;
      page++;
    }

    // sync_log 완료
    if (logEntry?.id) {
      await db.from('sync_log').update({
        status: 'completed',
        records_synced: synced,
        error_message: errors.length > 0 ? errors.join('\n') : null,
        completed_at: new Date().toISOString(),
      }).eq('id', logEntry.id);
    }

    return { success: true, synced, created, updated, errors };
  } catch (err) {
    if (logEntry?.id) {
      await db.from('sync_log').update({
        status: 'failed',
        records_synced: synced,
        error_message: String(err),
        completed_at: new Date().toISOString(),
      }).eq('id', logEntry.id);
    }

    return { success: false, synced, created, updated, errors: [String(err), ...errors] };
  }
}
