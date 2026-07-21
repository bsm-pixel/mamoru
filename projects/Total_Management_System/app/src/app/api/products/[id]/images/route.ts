import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';

/**
 * 재고판매(LS) 상세 이미지 관리 — 2026-07-21
 *
 * 여러 장의 상세 이미지를 products.tags.images (URL 배열) 에 저장한다.
 * 대표 썸네일(image_url) = images[0]. 스토리지는 기존 공개 버킷 repair-photos 재사용(products/ 경로).
 *
 * POST   multipart file(들) 업로드 → tags.images 뒤에 append
 * PATCH  { images: string[] } 순서 재정렬 (대표컷 = 첫 장)
 * DELETE ?url=... 1장 제거 (스토리지에서도 삭제)
 *
 * ⚠️ 재고 수량 로직 무관 — 이미지 URL 배열만 갱신.
 */

const BUCKET = 'repair-photos';
const MAX_SIZE = 10 * 1024 * 1024;
const MAX_IMAGES = 8;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function currentImages(db: any, id: string): Promise<{ tags: Record<string, unknown>; images: string[] }> {
  const { data, error } = await db.from('products').select('tags').eq('id', id).single();
  if (error) throw error;
  const tags = (data?.tags || {}) as Record<string, unknown>;
  const images = Array.isArray(tags.images) ? (tags.images as string[]) : [];
  return { tags, images };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function saveImages(db: any, id: string, tags: Record<string, unknown>, images: string[]) {
  await db.from('products').update({
    tags: { ...tags, images },
    image_url: images[0] || null,   // 대표 썸네일 = 첫 장
  }).eq('id', id);
}

/** 스토리지 공개 URL → 내부 경로 추출 (삭제용) */
function pathFromUrl(url: string): string | null {
  const m = url.match(/\/object\/public\/[^/]+\/(.+)$/);
  return m ? decodeURIComponent(m[1]) : null;
}

/** POST — 이미지 여러 장 업로드 */
export async function POST(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const authed = await createServerSupabaseClient();
    const { data: { user } } = await authed.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const form = await req.formData();
    const files = form.getAll('file').filter((f): f is File => f instanceof File);
    if (files.length === 0) return NextResponse.json({ error: 'file 필수' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { tags, images } = await currentImages(db, id);
    if (images.length + files.length > MAX_IMAGES) {
      return NextResponse.json({ error: `이미지는 최대 ${MAX_IMAGES}장입니다` }, { status: 400 });
    }

    const added: string[] = [];
    for (const file of files) {
      if (file.size > MAX_SIZE) return NextResponse.json({ error: '파일 10MB 초과' }, { status: 400 });
      const ext = (file.name.split('.').pop() || 'jpg').toLowerCase();
      const filePath = `products/${id}/${Date.now()}-${added.length}.${ext}`;
      const { error: upErr } = await db.storage.from(BUCKET).upload(filePath, file, { contentType: file.type, upsert: false });
      if (upErr) throw upErr;
      added.push(db.storage.from(BUCKET).getPublicUrl(filePath).data.publicUrl);
    }

    const next = [...images, ...added].slice(0, MAX_IMAGES);
    await saveImages(db, id, tags, next);
    return NextResponse.json({ images: next }, { status: 201 });
  } catch (err) {
    console.error('[products/images POST] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** PATCH — 순서 재정렬 (대표컷 지정 포함) */
export async function PATCH(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const authed = await createServerSupabaseClient();
    const { data: { user } } = await authed.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const body = await req.json();
    const order = Array.isArray(body.images) ? (body.images as string[]) : null;
    if (!order) return NextResponse.json({ error: 'images 배열 필요' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { tags, images } = await currentImages(db, id);
    // 기존 이미지 집합과 동일한 것만 허용(재정렬만, 새 URL 주입 금지)
    const set = new Set(images);
    const next = order.filter((u) => set.has(u));
    if (next.length !== images.length) return NextResponse.json({ error: '이미지 목록이 일치하지 않습니다' }, { status: 400 });
    await saveImages(db, id, tags, next);
    return NextResponse.json({ images: next });
  } catch (err) {
    console.error('[products/images PATCH] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

/** DELETE ?url=... — 1장 제거 */
export async function DELETE(req: NextRequest, { params }: { params: Promise<{ id: string }> }) {
  try {
    const { id } = await params;
    const authed = await createServerSupabaseClient();
    const { data: { user } } = await authed.auth.getUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

    const url = req.nextUrl.searchParams.get('url');
    if (!url) return NextResponse.json({ error: 'url 필요' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { tags, images } = await currentImages(db, id);
    const next = images.filter((u) => u !== url);
    await saveImages(db, id, tags, next);

    // 스토리지에서도 정리 (실패해도 무시 — URL 은 이미 제거됨)
    const path = pathFromUrl(url);
    if (path) { try { await db.storage.from(BUCKET).remove([path]); } catch { /* ignore */ } }

    return NextResponse.json({ images: next });
  } catch (err) {
    console.error('[products/images DELETE] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
