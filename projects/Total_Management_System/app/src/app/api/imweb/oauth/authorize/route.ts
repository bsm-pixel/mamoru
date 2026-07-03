import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/imweb/oauth/authorize — 아임웹 OpenAPI 재인증(재연결) 시작
 * 아임웹 OAuth authorize로 302 리다이렉트 → 승인 후 /api/imweb/oauth/callback 로 code 수신.
 *
 * refreshToken 만료/무효화로 재고 push(product:write)가 죽었을 때 토큰·권한을 재발급받는 유일한 경로.
 * scope는 활성 전체(site-info/product/script/order:write)를 함께 재발급해 모든 연동을 한 번에 복구.
 */
const SITE_CODE = 'S20250825bc9b09c7146df';
const SCOPES = ['site-info:write', 'product:write', 'script:write', 'order:write'];

export async function GET(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const clientId = process.env.IMWEB_OPENAPI_KEY;
  if (!clientId) {
    return NextResponse.json({ error: 'IMWEB_OPENAPI_KEY 환경변수가 없습니다' }, { status: 500 });
  }

  const redirectUri = `${req.nextUrl.origin}/api/imweb/oauth/callback`;
  const query = [
    'responseType=code',
    `clientId=${encodeURIComponent(clientId)}`,
    `redirectUri=${encodeURIComponent(redirectUri)}`,
    `scope=${encodeURIComponent(SCOPES.join(' '))}`,
    `siteCode=${SITE_CODE}`,
  ].join('&');

  return NextResponse.redirect(`https://openapi.imweb.me/oauth2/authorize?${query}`);
}
