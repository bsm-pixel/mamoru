import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

/** 복원수리 진행방식(proceed_type) → 리뷰 subtype (후기 칩 방문수거/직접발송/직접방문 구분) */
function proceedToSubtype(pt: string | null | undefined): string {
  if (pt === '방문수거') return 'parcel_pickup';
  if (pt === '직접발송') return 'self_ship';
  if (pt === '직접방문') return 'direct_visit';
  return 'restoration';
}

/** 리뷰 ID 자동 채번: RV-YYYYMMDD-NNN */
async function generateReviewId(db: ReturnType<typeof createServiceClient>): Promise<string> {
  const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const prefix = `RV-${today}-`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const { data } = await (db as any)
    .from('reviews')
    .select('review_id')
    .like('review_id', `${prefix}%`)
    .order('review_id', { ascending: false })
    .limit(1);

  let seq = 1;
  if (data && data.length > 0) {
    const last = data[0].review_id as string;
    seq = parseInt(last.split('-').pop() || '0', 10) + 1;
  }

  return `${prefix}${String(seq).padStart(3, '0')}`;
}

/** POST /api/reviews/submit — 고객 리뷰 제출 (비인증, 공개 엔드포인트) */
export async function POST(req: NextRequest) {
  try {
    const body = await req.json();
    const { uid, type: rawType, stars, content, photoUrls, productNo, tags, subtype: bodySubtype } = body;
    // 'as'는 'repair'의 별칭으로 정규화 (info route와 동일 패턴)
    const type = rawType === 'as' ? 'repair' : rawType;

    if (!uid || !type || !stars || !content) {
      return NextResponse.json(
        { error: '필수 항목을 입력해주세요' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    if (stars < 1 || stars > 5) {
      return NextResponse.json(
        { error: '별점은 1~5 사이여야 합니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    // Service Role 클라이언트 (RLS 우회하여 조회 가능)
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 중복 제출 방지: purchase는 제품별 source_id, 나머지는 uid
    const sourceId = type === 'purchase' ? `${uid}:${productNo}` : uid;
    const { data: existing } = await dbAny
      .from('reviews')
      .select('id')
      .eq('source_id', sourceId)
      .limit(1);

    if (existing && existing.length > 0) {
      return NextResponse.json(
        { error: '이미 후기를 작성하셨습니다' },
        { status: 409, headers: CORS_HEADERS }
      );
    }

    // uid로 원본 데이터 조회 (상담 또는 수리)
    let name = '';
    let phone = '';
    let subtype = '';
    let meta: Record<string, string> = {};

    if (type === 'consult') {
      const { data: consult } = await dbAny
        .from('consultations')
        .select('name, phone, consultation_type, visit_date, visit_time, created_at')
        .eq('unique_id', uid)
        .single();

      if (consult) {
        name = consult.name;
        phone = consult.phone || '';
        subtype = consult.consultation_type || '';
        meta = {
          visit_date: consult.visit_date || '',
          visit_time: consult.visit_time || '',
          consultation_type: consult.consultation_type || '',
          received_at: consult.created_at || '',
        };
      } else {
        // 판매 건(OS-*) fallback: offline_sales에서 조회
        const { data: sale } = await dbAny
          .from('offline_sales')
          .select('customer_name, customer_phone, sale_channel, sale_date, created_at, review_promised_subtype')
          .eq('sale_number', uid)
          .single();

        if (!sale) {
          return NextResponse.json(
            { error: '상담 정보를 찾을 수 없습니다' },
            { status: 404, headers: CORS_HEADERS }
          );
        }
        name = sale.customer_name;
        phone = sale.customer_phone || '';
        // 판매건에 건 리뷰약속의 promised subtype 우선 (없을 때만 sale_channel fallback)
        // → 매장방문(store_visit) 상담이 'offline'(sale_channel)로 잘못 박히는 문제 방지
        subtype = bodySubtype || sale.review_promised_subtype || sale.sale_channel || '';
        meta = {
          sale_number: uid,
          sale_channel: sale.sale_channel || '',
          sale_date: sale.sale_date || '',
          received_at: sale.created_at || '',
        };
      }
    } else if (type === 'repair') {
      const { data: repair } = await dbAny
        .from('repairs')
        .select('name, phone, proceed_type, created_at')
        .eq('as_id', uid)
        .single();

      if (repair) {
        name = repair.name;
        phone = repair.phone || '';
        // proceed_type → subtype 파생(방문수거/직접발송/직접방문 구분). bodySubtype(약속/URL) 우선
        subtype = bodySubtype || proceedToSubtype(repair.proceed_type);
        meta = {
          proceed_type: repair.proceed_type || '',
          received_at: repair.created_at || '',
        };
      } else {
        // 판매 건(OS-*) fallback: offline_sales에서 조회
        const { data: sale } = await dbAny
          .from('offline_sales')
          .select('customer_name, customer_phone, sale_channel, sale_date, created_at, review_promised_subtype')
          .eq('sale_number', uid)
          .single();

        if (!sale) {
          return NextResponse.json(
            { error: '복원수리 정보를 찾을 수 없습니다' },
            { status: 404, headers: CORS_HEADERS }
          );
        }
        name = sale.customer_name;
        phone = sale.customer_phone || '';
        subtype = bodySubtype || sale.review_promised_subtype || 'restoration';
        meta = {
          sale_number: uid,
          sale_channel: sale.sale_channel || '',
          sale_date: sale.sale_date || '',
          received_at: sale.created_at || '',
        };
      }
    } else if (type === 'purchase') {
      if (!productNo) {
        return NextResponse.json(
          { error: '제품을 선택해주세요' },
          { status: 400, headers: CORS_HEADERS }
        );
      }
      // (1) 아임웹 온라인 주문 (id=uid)
      const { data: order } = await dbAny
        .from('orders')
        .select('orderer_name, orderer_phone, ordered_at')
        .eq('id', uid)
        .single();

      if (order) {
        const { data: item } = await dbAny
          .from('order_items')
          .select('product_name, imweb_product_no')
          .eq('order_id', uid)
          .eq('imweb_product_no', productNo)
          .limit(1)
          .single();

        name = order.orderer_name;
        phone = order.orderer_phone || '';
        subtype = '';
        meta = {
          order_id: uid,
          imweb_product_no: productNo,
          product_name: item?.product_name || '',
          received_at: order.ordered_at || '',
        };
        if (productNo) {
          const { data: prod } = await dbAny
            .from('products').select('product_group').eq('imweb_product_no', productNo).single();
          if (prod?.product_group) meta.product_group = prod.product_group;
        }
      } else {
        // (2) 오프라인 판매(OS-) fallback — 재고판매·매장 등 offline_sales 기반 제품구매 후기
        const { data: sale } = await dbAny
          .from('offline_sales')
          .select('id, customer_name, customer_phone, sale_date, created_at, memo')
          .eq('sale_number', uid)
          .single();
        if (!sale) {
          return NextResponse.json({ error: '주문 정보를 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
        }
        // productNo(키) 가 imweb_product_no 이거나 product_id 일 수 있다 → 판매 품목에서 매칭
        const { data: sitems } = await dbAny
          .from('offline_sale_items').select('product_id, product_name').eq('sale_id', sale.id);
        const pids = [...new Set((sitems || []).map((s: { product_id: string | null }) => s.product_id).filter(Boolean))] as string[];
        const prodMap = new Map<string, { imweb_product_no: string | null; product_group: string | null }>();
        if (pids.length > 0) {
          const { data: prods } = await dbAny.from('products').select('id, imweb_product_no, product_group').in('id', pids);
          (prods || []).forEach((p: { id: string; imweb_product_no: string | null; product_group: string | null }) =>
            prodMap.set(p.id, { imweb_product_no: p.imweb_product_no ? String(p.imweb_product_no) : null, product_group: p.product_group }));
        }
        // 키가 일치하는 품목 찾기 (imweb 번호 우선, 없으면 product_id)
        const matched = (sitems || []).find((s: { product_id: string | null }) => {
          const pm = s.product_id ? prodMap.get(s.product_id) : undefined;
          const key = (pm?.imweb_product_no) || s.product_id || '';
          return key === String(productNo);
        }) as { product_id: string | null; product_name: string } | undefined;
        if (!matched) {
          return NextResponse.json({ error: '제품을 찾을 수 없습니다' }, { status: 404, headers: CORS_HEADERS });
        }
        const pm = matched.product_id ? prodMap.get(matched.product_id) : undefined;
        name = sale.customer_name;
        phone = sale.customer_phone || '';
        // EVENT 전환 판매(memo 'EVENT 전환'…)에서 온 구매 후기 → subtype='event' (리뷰관리/후기벽 '이벤트' 칩)
        subtype = (typeof sale.memo === 'string' && sale.memo.startsWith('EVENT 전환')) ? 'event' : '';
        meta = {
          sale_number: uid,
          imweb_product_no: pm?.imweb_product_no || '',   // 실제 아임웹 번호일 때만 (이미지 연동용)
          product_name: matched.product_name || '',
          received_at: sale.created_at || sale.sale_date || '',
        };
        if (pm?.product_group) meta.product_group = pm.product_group;
      }
    } else {
      return NextResponse.json(
        { error: '지원하지 않는 리뷰 유형입니다' },
        { status: 400, headers: CORS_HEADERS }
      );
    }

    const reviewId = await generateReviewId(db);

    // 자동 노출 설정 체크
    let reviewStatus = 'pending';
    const { data: autoSetting } = await dbAny
      .from('system_settings')
      .select('value')
      .eq('key', 'review.auto_approve')
      .single();
    if (autoSetting && autoSetting.value === 'true') {
      reviewStatus = 'approved';
    }

    const { data, error } = await dbAny
      .from('reviews')
      .insert({
        review_id: reviewId,
        type,
        subtype: subtype || null,
        name,
        phone,
        stars: Number(stars),
        content: String(content).trim(),
        photo_urls: Array.isArray(photoUrls) ? photoUrls : [],
        source_id: sourceId,
        product: meta.product_name || null,
        product_group: meta.product_group || null,
        status: reviewStatus,
        approved_at: reviewStatus === 'approved' ? new Date().toISOString() : null,
        meta: { ...meta, ...(Array.isArray(tags) && tags.length > 0 ? { tags } : {}) },
      })
      .select()
      .single();

    if (error) throw error;

    // 067: 역방향 매칭 — source 테이블의 review_submitted_at 자동 set
    // - consult: consultations.unique_id 또는 offline_sales.sale_number (talk 채널 fallback)
    // - repair:  repairs.as_id 또는 offline_sales.sale_number (repair 채널 fallback)
    // .is('review_submitted_at', null)로 멱등성 보장 + service role이라 RLS 무관
    after(async () => {
      try {
        const submittedAt = new Date().toISOString();
        if (type === 'consult') {
          await dbAny.from('consultations')
            .update({ review_submitted_at: submittedAt })
            .eq('unique_id', sourceId)
            .is('review_submitted_at', null);
          await dbAny.from('offline_sales')
            .update({ review_submitted_at: submittedAt })
            .eq('sale_number', sourceId)
            .is('review_submitted_at', null);
        } else if (type === 'repair') {
          await dbAny.from('repairs')
            .update({ review_submitted_at: submittedAt })
            .eq('as_id', sourceId)
            .is('review_submitted_at', null);
          await dbAny.from('offline_sales')
            .update({ review_submitted_at: submittedAt })
            .eq('sale_number', sourceId)
            .is('review_submitted_at', null);
        } else if (type === 'purchase') {
          // 제품구매 후기는 제품별(source_id=`uid:key`)이라 uid 는 sourceId 가 아닌 원본 uid 를 써야 한다.
          // 오프라인 판매(OS-)면 offline_sales 에 표시 (아임웹 orders 는 review_submitted_at 컬럼 없음 — 무시).
          await dbAny.from('offline_sales')
            .update({ review_submitted_at: submittedAt })
            .eq('sale_number', uid)
            .is('review_submitted_at', null);
        }
      } catch (e) {
        console.error('[reviews/submit reverse-match] 실패:', e);
      }
    });

    // 관리자 푸시 알림 — after()로 응답 후 실행 보장 (Vercel 서버리스 대응)
    after(async () => {
      try {
        const { sendPushToAll } = await import('@/lib/firebase/send-push');
        const typeLabel = type === 'repair' ? '복원수리' : type === 'consult' ? '상담' : '제품구매';
        await sendPushToAll({
          title: `새 리뷰 도착 ⭐${stars}`,
          body: `${name}님 ${typeLabel} 리뷰 — ${String(content).slice(0, 40)}${String(content).length > 40 ? '...' : ''}`,
          url: '/reviews',
          tag: 'mamoru-review',
          settingKey: 'push.review_submitted',
        });
      } catch (e) {
        console.error('[reviews/submit push] 실패:', e);
      }
    });

    return NextResponse.json(
      { success: true, reviewId: data.review_id },
      { headers: CORS_HEADERS }
    );
  } catch (err) {
    console.error('[reviews/submit] 제출 실패:', err);
    return NextResponse.json(
      { error: String(err) },
      { status: 500, headers: CORS_HEADERS }
    );
  }
}
