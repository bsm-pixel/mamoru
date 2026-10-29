import { NextRequest, NextResponse } from 'next/server';
import { requireUser } from '@/lib/auth/require-user';
import { createServiceClient } from '@/lib/supabase/server';
import { RECEIPT_BUCKET, receiptPaths } from '@/lib/expenses/receipt-archive';

/**
 * GET /api/expenses/receipt?id=<경비 id>&i=<몇 번째 파일, 기본 0> — 보관한 영수증 원본 열기 (2026-10-29)
 * 로그인 확인 → PDF 는 1분짜리 서명 주소로 이동, 메일 본문(.html)은 직접 보여 줌(sandbox). 간편장부 엑셀의 링크도 여기로 온다(엑셀은 영구 주소, 열 때마다 새 서명).
 */
export async function GET(req: NextRequest) {
  const deny = await requireUser(); if (deny) return deny;
  const id = req.nextUrl.searchParams.get('id');
  const i = parseInt(req.nextUrl.searchParams.get('i') || '0', 10) || 0;
  if (!id) return NextResponse.json({ error: 'id 필수' }, { status: 400 });
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const db = createServiceClient() as any;
  const { data: e } = await db.from('expenses').select('receipt_url').eq('id', id).single();
  const path = receiptPaths(e?.receipt_url)[i];
  if (!path) return NextResponse.json({ error: '보관된 영수증이 없습니다' }, { status: 404 });
  // 메일 본문(.html)은 저장소가 글자(text/plain)로 내줘 코드가 보인다 → TMS 가 직접 내주되 sandbox 로 스크립트·쿠키 접근 차단
  if (path.endsWith('.html')) {
    const { data: file, error: dlErr } = await db.storage.from(RECEIPT_BUCKET).download(path);
    if (dlErr || !file) return NextResponse.json({ error: String(dlErr?.message || '파일 없음') }, { status: 500 });
    return new NextResponse(await file.text(), {
      headers: {
        'Content-Type': 'text/html; charset=utf-8',
        'Content-Security-Policy': "sandbox; default-src 'none'; img-src * data:; style-src 'unsafe-inline' *; font-src * data:",
        'X-Content-Type-Options': 'nosniff',
        'Cache-Control': 'private, no-store',
      },
    });
  }
  const { data, error } = await db.storage.from(RECEIPT_BUCKET).createSignedUrl(path, 60);
  if (error || !data?.signedUrl) return NextResponse.json({ error: String(error?.message || '주소 생성 실패') }, { status: 500 });
  return NextResponse.redirect(data.signedUrl);
}
