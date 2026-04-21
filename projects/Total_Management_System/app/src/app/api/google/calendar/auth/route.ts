/**
 * GET /api/google/calendar/auth
 * OAuth 2.0 인가 URL 생성 → Google 로그인 페이지로 리다이렉트
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { createOAuthClient, SCOPES } from '@/lib/google/oauth';

export async function GET(req: NextRequest) {
  // 관리자 인증 필요
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    return NextResponse.redirect(new URL('/login', req.nextUrl.origin));
  }

  try {
    const oauth2Client = createOAuthClient();
    const url = oauth2Client.generateAuthUrl({
      access_type: 'offline', // refresh_token 발급 필수
      prompt: 'consent', // 재연결 시에도 refresh_token 강제 재발급
      scope: SCOPES,
      include_granted_scopes: true,
    });

    return NextResponse.redirect(url);
  } catch (err) {
    return NextResponse.redirect(
      new URL(
        `/settings?google_calendar=error&msg=${encodeURIComponent(String(err))}`,
        req.nextUrl.origin,
      ),
    );
  }
}
