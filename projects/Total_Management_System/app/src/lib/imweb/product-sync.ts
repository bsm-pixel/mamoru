/**
 * 아임웹 상품 동기화
 * 아임웹 OpenAPI → TMS products 테이블 upsert
 * 매입가/거래처는 동기화하지 않음 (TMS 자체 관리)
 */

import { getImwebProducts } from './client';
import { createServiceClient } from '@/lib/supabase/server';

interface SyncResult {
  success: boolean;
  synced: number;
  created: number;
  updated: number;
  errors: string[];
}

/** 아임웹 카테고리 → TMS 카테고리 매핑 (기본값) */
function mapCategory(categories: Array<{ categoryName: string }>): string {
  const name = categories?.[0]?.categoryName?.toLowerCase() || '';
  if (name.includes('블런트') || name.includes('blunt')) return 'BL';
  if (name.includes('틴닝') || name.includes('thinning')) return 'TH';
  if (name.includes('장가위') || name.includes('long')) return 'LO';
  if (name.includes('슬라이싱') || name.includes('slicing')) return 'SL';
  return 'BL'; // 기본값
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
      const res = await getImwebProducts(page, 100);

      if (res.statusCode !== 200 || !res.data?.list) {
        errors.push(`페이지 ${page} 조회 실패: statusCode ${res.statusCode}`);
        break;
      }

      const products = res.data.list;

      for (const imwebProd of products) {
        try {
          // 기존 TMS 제품 조회 (imweb_product_no 기준)
          const { data: existing } = await db
            .from('products')
            .select('id, price_purchase, price_dealer, supplier_id, category')
            .eq('imweb_product_no', String(imwebProd.prodNo))
            .single();

          const productData = {
            name: imwebProd.name,
            imweb_product_no: String(imwebProd.prodNo),
            price: imwebProd.price || 0,
            stock_quantity: imwebProd.stock || 0,
            image_url: imwebProd.imageUrl || null,
            sku: imwebProd.customSkuCode || `IW-${imwebProd.prodNo}`,
            is_active: imwebProd.prodStatus === 'sale',
            updated_at: new Date().toISOString(),
          };

          if (existing) {
            // 업데이트 — 매입가/거래처/카테고리는 TMS 값 유지
            await db
              .from('products')
              .update(productData)
              .eq('id', existing.id);
            updated++;
          } else {
            // 신규 생성
            await db.from('products').insert({
              ...productData,
              category: mapCategory(imwebProd.categories || []),
              price_dealer: 0,
              price_purchase: 0,
              created_at: new Date().toISOString(),
            });
            created++;
          }
          synced++;
        } catch (err) {
          errors.push(`상품 ${imwebProd.prodNo}(${imwebProd.name}): ${String(err)}`);
        }
      }

      // 페이지네이션
      hasMore = page < res.data.lastPage;
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
    // sync_log 실패
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
