import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

/**
 * GET /api/imweb/oauth/callback — 아임웹 OpenAPI OAuth2 콜백
 * 인가 코드(code)를 받아서 accessToken + refreshToken을 발급받고 DB에 저장
 */
export async function GET(req: NextRequest) {
  try {
    const code = req.nextUrl.searchParams.get('code');
    if (!code) {
      return NextResponse.json({ error: 'code 파라미터가 없습니다' }, { status: 400 });
    }

    const clientId = process.env.IMWEB_OPENAPI_KEY;
    const clientSecret = process.env.IMWEB_OPENAPI_SECRET;
    const redirectUri = `${req.nextUrl.origin}/api/imweb/oauth/callback`;

    // Access Token 발급
    const tokenRes = await fetch('https://openapi.imweb.me/oauth2/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        clientId: clientId || '',
        clientSecret: clientSecret || '',
        code,
        grantType: 'authorization_code',
        redirectUri,
      }),
    });

    const tokenData = await tokenRes.json();

    if (!tokenRes.ok || tokenData.statusCode !== 200) {
      return NextResponse.json({
        error: '토큰 발급 실패',
        detail: tokenData,
      }, { status: 400 });
    }

    const { accessToken, refreshToken, scope } = tokenData.data;

    // DB에 토큰 저장 (system_settings)
    const supabase = await createServerSupabaseClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = supabase as any;

    const now = new Date().toISOString();
    const tokenSettings = [
      { key: 'imweb_openapi.access_token', value: accessToken },
      { key: 'imweb_openapi.refresh_token', value: refreshToken },
      { key: 'imweb_openapi.scope', value: scope || '' },
      { key: 'imweb_openapi.token_updated_at', value: now },
    ];

    for (const { key, value } of tokenSettings) {
      await db.from('system_settings').upsert(
        { key, value, updated_at: now },
        { onConflict: 'key' },
      );
    }

    // 성공 페이지로 리다이렉트
    return NextResponse.redirect(new URL('/settings?imweb_oauth=success', req.nextUrl.origin));
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
