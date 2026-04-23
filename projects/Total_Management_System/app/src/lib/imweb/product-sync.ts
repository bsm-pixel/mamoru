/**
 * 아임웹 상품 동기화 (v2 API)
 * 아임웹 → TMS products 테이블 upsert
 * 매입가/거래처는 동기화하지 않음 (TMS 자체 관리)
 *
 * 매칭 우선순위:
 *   1. imweb_product_no 로 기존 조회 (이미 연동된 상품)
 *   2. sku (custom_prod_code) 로 기존 조회 (수동 등록된 상품)
 *   3. 없으면 신규 생성
 *
 * 에러 처리: 각 상품별 INSERT/UPDATE 실패 사유를 수집해 반환
 */

import { getImwebProducts, type ImwebV2Product } from './client';
import { createServiceClient } from '@/lib/supabase/server';

interface SyncResult {
  success: boolean;
  total_fetched: number;   // 아임웹 API에서 받은 전체 개수
  synced: number;          // 실제 DB에 upsert 성공한 개수
  created: number;
  updated: number;
  linked: number;          // 수동 등록된 것에 imweb_product_no 연결
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
  let totalFetched = 0;
  let synced = 0;
  let created = 0;
  let updated = 0;
  let linked = 0;

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
      totalFetched += products.length;

      for (const p of products) {
        try {
          const sku = p.custom_prod_code || `IW-${p.no}`;

          // 1차: imweb_product_no 로 기존 조회
          const { data: byImwebNo, error: imwebErr } = await db
            .from('products')
            .select('id, sku, imweb_product_no')
            .eq('imweb_product_no', String(p.no))
            .maybeSingle();

          if (imwebErr) {
            errors.push(`[${p.no}] ${p.name}: imweb_product_no 조회 오류 — ${imwebErr.message}`);
            continue;
          }

          // 재고 미사용: stock_use=false → stock_quantity=-1 (TMS 규약)
          const stockQty = p.stock?.stock_use ? (p.stock.stock_no_option || 0) : -1;

          const productData = {
            name: p.name,
            imweb_product_no: String(p.no),
            price: p.price || 0,
            stock_quantity: stockQty,
            image_url: getImageUrl(p),
            sku,
            is_active: p.prod_status === 'sale',
            updated_at: new Date().toISOString(),
          };

          if (byImwebNo) {
            // 이미 연동된 상품 → UPDATE (재고는 TMS가 마스터라 제외)
            const { stock_quantity: _omit, ...updateData } = productData;
            const { error: updErr } = await db
              .from('products')
              .update(updateData)
              .eq('id', byImwebNo.id);

            if (updErr) {
              errors.push(`[${p.no}] ${p.name}: UPDATE 실패 — ${updErr.message}`);
              continue;
            }
            updated++;
          } else {
            // 2차: sku 로 수동 등록된 상품 조회 (imweb_product_no 아직 안 붙은 상품)
            const { data: bySku, error: skuErr } = await db
              .from('products')
              .select('id, sku, imweb_product_no')
              .eq('sku', sku)
              .maybeSingle();

            if (skuErr) {
              errors.push(`[${p.no}] ${p.name} (sku=${sku}): sku 조회 오류 — ${skuErr.message}`);
              continue;
            }

            if (bySku) {
              // 수동 등록된 상품을 이번 동기화로 imweb와 연결
              const { stock_quantity: _omit, ...updateData } = productData;
              const { error: linkErr } = await db
                .from('products')
                .update(updateData)
                .eq('id', bySku.id);

              if (linkErr) {
                errors.push(`[${p.no}] ${p.name} (sku=${sku}): 연결(link) 실패 — ${linkErr.message}`);
                continue;
              }
              linked++;
            } else {
              // 신규 생성 — 아임웹 재고를 보관창고(raw_stock)에 초기화
              const { error: insErr } = await db.from('products').insert({
                ...productData,
                category: 'BL', // 기본값, TMS에서 수동 변경
                price_dealer: 0,
                price_purchase: 0,
                price_groups: {},
                raw_stock: stockQty > 0 ? stockQty : 0, // 아임웹 재고 → 보관창고
                created_at: new Date().toISOString(),
              });

              if (insErr) {
                errors.push(`[${p.no}] ${p.name} (sku=${sku}): INSERT 실패 — ${insErr.message}`);
                continue;
              }
              created++;
            }
          }
          synced++;
        } catch (err) {
          errors.push(`[${p.no}] ${p.name}: 예외 — ${String(err)}`);
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

    // 콘솔에 결과 요약 (Vercel 로그 확인용)
    console.log(
      `[syncProducts] fetched=${totalFetched} synced=${synced} created=${created} updated=${updated} linked=${linked} errors=${errors.length}`
    );
    if (errors.length > 0) {
      console.warn('[syncProducts] 실패 상세:\n' + errors.slice(0, 10).join('\n'));
    }

    return { success: true, total_fetched: totalFetched, synced, created, updated, linked, errors };
  } catch (err) {
    if (logEntry?.id) {
      await db.from('sync_log').update({
        status: 'failed',
        records_synced: synced,
        error_message: String(err),
        completed_at: new Date().toISOString(),
      }).eq('id', logEntry.id);
    }

    return { success: false, total_fetched: totalFetched, synced, created, updated, linked, errors: [String(err), ...errors] };
  }
}
