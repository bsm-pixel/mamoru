/**
 * GET /api/google/calendar/callback
 * Google OAuth 콜백 — 인가 code를 토큰으로 교환하고 DB에 저장
 * id_token에서 email/hd 추출해 Workspace vs 일반 Gmail 자동 판별
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { createOAuthClient } from '@/lib/google/oauth';

interface IdTokenPayload {
  email?: string;
  hd?: string; // Workspace 도메인 (일반 Gmail은 없음)
  email_verified?: boolean;
}

export async function GET(req: NextRequest) {
  const origin = req.nextUrl.origin;

  // 관리자 인증 필요
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) {
    return NextResponse.redirect(new URL('/login', origin));
  }

  const code = req.nextUrl.searchParams.get('code');
  const oauthError = req.nextUrl.searchParams.get('error');

  if (oauthError || !code) {
    return NextResponse.redirect(
      new URL(
        `/settings?google_calendar=error&msg=${encodeURIComponent(oauthError || 'no_code')}`,
        origin,
      ),
    );
  }

  try {
    const oauth2Client = createOAuthClient();
    const { tokens } = await oauth2Client.getToken(code);

    if (!tokens.refresh_token) {
      return NextResponse.redirect(
        new URL(
          `/settings?google_calendar=error&msg=${encodeURIComponent('refresh_token_missing — Google Cloud Console에서 앱을 완전히 재승인해주세요')}`,
          origin,
        ),
      );
    }

    // id_token 디코드 → email / hd 추출
    let email = '';
    let hd = '';
    if (tokens.id_token) {
      const parts = tokens.id_token.split('.');
      if (parts.length === 3) {
        try {
          // base64url → base64
          const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
          const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4);
          const payload = JSON.parse(Buffer.from(padded, 'base64').toString('utf-8')) as IdTokenPayload;
          email = payload.email || '';
          hd = payload.hd || '';
        } catch {
          /* id_token 디코드 실패 — 핵심 기능 아니므로 무시 */
        }
      }
    }

    // system_settings UPSERT
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;
    const now = new Date().toISOString();

    const upserts: Array<{ key: string; value: string }> = [
      { key: 'google.calendar.access_token', value: tokens.access_token || '' },
      { key: 'google.calendar.refresh_token', value: tokens.refresh_token },
      {
        key: 'google.calendar.token_expires_at',
        value: tokens.expiry_date ? new Date(tokens.expiry_date).toISOString() : '',
      },
      { key: 'google.calendar.connected_email', value: email },
      { key: 'google.calendar.connected_hd', value: hd },
      { key: 'google.calendar.connected_at', value: now },
      { key: 'google.calendar.calendar_id', value: 'primary' },
    ];

    for (const { key, value } of upserts) {
      if (value) {
        await dbAny.from('system_settings').upsert(
          { key, value, updated_at: now },
          { onConflict: 'key' },
        );
      }
    }

    // 이전 오류 지우기
    await dbAny.from('system_settings').delete().eq('key', 'google.calendar.last_error');

    return NextResponse.redirect(new URL('/settings?google_calendar=connected', origin));
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return NextResponse.redirect(
      new URL(`/settings?google_calendar=error&msg=${encodeURIComponent(msg)}`, origin),
    );
  }
}
