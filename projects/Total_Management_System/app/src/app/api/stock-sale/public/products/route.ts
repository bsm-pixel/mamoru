import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

/** 재고판매 재고 라벨 — 범용 문구('개'). EVENT의 stockLabel(자루·획득가능)은 가위 전용이라 별도 */
function lsStockLabel(stock: number): { text: string; tone: 'soldout' | 'low' | 'ok' } {
  if (stock <= 0) return { text: '품절', tone: 'soldout' };
  if (stock <= 5) return { text: `${stock}개 남음`, tone: 'low' };
  return { text: '재고 있음', tone: 'ok' };
}

/**
 * 재고판매(LS) 공개 카탈로그 API — 비인증 + CORS (2026-07-21)
 *
 * 고객 대면 카탈로그 폼(page.mamoru.kr)이 fetch 한다.
 * category='LS' + is_active 인 제품을 사진·가격·재고와 함께 반환.
 * EVENT 카탈로그(api/event/public/products)와 형제 — 가위 옵션(손/종류/spec) 없이 범용.
 *
 * ⚠️ 재고 수량은 읽기만. 주문 접수는 submit API 에서, 재고 차감은 판매전환에서 처리.
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** GET /api/stock-sale/public/products?include_soldout=1 */
export async function GET(req: Request) {
  try {
    const url = new URL(req.url);
    const includeSoldout = url.searchParams.get('include_soldout') === '1';
    const db = createServiceClient();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const { data, error } = await (db as any)
      .from('products')
      .select('id, sku, name, price, stock_quantity, image_url, description, tags, is_active')
      .eq('category', 'LS')
      .eq('is_active', true)
      .order('name', { ascending: true });

    if (error) throw error;

    const items = (data || [])
      .map((p: {
        id: string; sku: string; name: string; price: number | null;
        stock_quantity: number | null; image_url: string | null; description: string | null;
        tags: Record<string, unknown> | null;
      }) => {
        const stock = p.stock_quantity ?? 0;
        const label = lsStockLabel(stock);
        const tags = p.tags || {};
        // 상세 이미지 여러 장 (tags.images) → 없으면 대표 썸네일 1장으로 대체
        const gallery = Array.isArray(tags.images) ? (tags.images as string[]).filter(Boolean) : [];
        const images = gallery.length > 0 ? gallery : (p.image_url ? [p.image_url] : []);
        return {
          id: p.id,
          sku: p.sku,
          name: p.name,
          price: p.price ?? 0,
          stock,                        // 담기 수량 상한 (표시는 stock_text 로 분기)
          soldout: stock <= 0,
          stock_text: label.text,
          stock_tone: label.tone,
          image_url: images[0] || '',
          images,                       // 상세 모달용 갤러리
          description: p.description || '',
          // 카탈로그 그룹핑용(선택) — 등록 시 tags.group 넣으면 폼에서 섹션 분리 가능
          group: (tags.group as string) || '',
        };
      })
      // 재고 없는 품목은 기본 숨김 (include_soldout=1 이면 회색으로 노출)
      .filter((it: { soldout: boolean }) => includeSoldout || !it.soldout);

    return NextResponse.json({ ok: true, items }, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[stock-sale/public/products] 실패:', err);
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
