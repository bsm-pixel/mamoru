/**
 * Google OAuth 2.0 클라이언트 — Calendar API 용
 * refresh_token은 system_settings에 저장 / access_token은 자동 갱신
 */

import { google } from 'googleapis';
import { createServiceClient } from '@/lib/supabase/server';

// googleapis가 내부적으로 두 개의 google-auth-library 버전을 포함하는 이슈 회피 위해
// google.auth.OAuth2 타입을 직접 추출
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type OAuth2ClientType = any;

interface GoogleCredentials {
  access_token?: string | null;
  refresh_token?: string | null;
  expiry_date?: number | null;
  id_token?: string | null;
}

const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID || '';
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET || '';
const GOOGLE_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI || '';

// openid email profile — id_token 발급용 (connected_email 자동 판별)
// calendar.events — 이벤트 CRUD 권한
export const SCOPES = [
  'openid',
  'email',
  'profile',
  'https://www.googleapis.com/auth/calendar.events',
];

/** 새 OAuth2 클라이언트 (토큰 없음 — 인가 단계에서 사용) */
export function createOAuthClient(): OAuth2ClientType {
  return new google.auth.OAuth2(GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REDIRECT_URI);
}

/** system_settings에서 토큰 로드 → refresh_token 기반 인증된 클라이언트 반환 */
export async function getAuthorizedClient(): Promise<OAuth2ClientType | null> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const { data } = await dbAny
    .from('system_settings')
    .select('key, value')
    .in('key', [
      'google.calendar.refresh_token',
      'google.calendar.access_token',
      'google.calendar.token_expires_at',
    ]);

  const map: Record<string, string> = {};
  (data || []).forEach((r: { key: string; value: string | null }) => {
    if (r.value) map[r.key] = String(r.value).replace(/^"|"$/g, '');
  });

  const refreshToken = map['google.calendar.refresh_token'];
  if (!refreshToken) return null;

  const client = createOAuthClient();
  client.setCredentials({
    refresh_token: refreshToken,
    access_token: map['google.calendar.access_token'] || undefined,
    expiry_date: map['google.calendar.token_expires_at']
      ? new Date(map['google.calendar.token_expires_at']).getTime()
      : undefined,
  });

  // 자동 갱신 이벤트 — 새 access_token 발급 시 DB에 저장
  client.on('tokens', (tokens: GoogleCredentials) => {
    saveTokensToDb(tokens).catch((err: unknown) => console.error('[google oauth] 토큰 저장 실패', err));
  });

  return client;
}

/** 토큰 DB 저장 (UPSERT) */
export async function saveTokensToDb(tokens: GoogleCredentials): Promise<void> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const upserts: Array<{ key: string; value: string }> = [];

  if (tokens.access_token) {
    upserts.push({ key: 'google.calendar.access_token', value: tokens.access_token });
  }
  if (tokens.refresh_token) {
    upserts.push({ key: 'google.calendar.refresh_token', value: tokens.refresh_token });
  }
  if (tokens.expiry_date) {
    upserts.push({
      key: 'google.calendar.token_expires_at',
      value: new Date(tokens.expiry_date).toISOString(),
    });
  }

  for (const u of upserts) {
    await dbAny.from('system_settings').upsert({ key: u.key, value: u.value }, { onConflict: 'key' });
  }
}

/** 연결 상태 조회 */
export async function getConnectionStatus(): Promise<{
  connected: boolean;
  email?: string;
  hd?: string;
  connected_at?: string;
  last_error?: string;
  last_success_at?: string;
}> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  const { data } = await dbAny
    .from('system_settings')
    .select('key, value')
    .in('key', [
      'google.calendar.refresh_token',
      'google.calendar.connected_email',
      'google.calendar.connected_hd',
      'google.calendar.connected_at',
      'google.calendar.last_error',
      'google.calendar.last_success_at',
    ]);

  const map: Record<string, string> = {};
  (data || []).forEach((r: { key: string; value: string | null }) => {
    if (r.value) map[r.key] = String(r.value).replace(/^"|"$/g, '');
  });

  return {
    connected: !!map['google.calendar.refresh_token'],
    email: map['google.calendar.connected_email'] || undefined,
    hd: map['google.calendar.connected_hd'] || undefined,
    connected_at: map['google.calendar.connected_at'] || undefined,
    last_error: map['google.calendar.last_error'] || undefined,
    last_success_at: map['google.calendar.last_success_at'] || undefined,
  };
}

/** 연결 해제 — 토큰 삭제 */
export async function disconnectCalendar(): Promise<void> {
  const db = createServiceClient();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const dbAny = db as any;

  await dbAny
    .from('system_settings')
    .delete()
    .in('key', [
      'google.calendar.access_token',
      'google.calendar.refresh_token',
      'google.calendar.token_expires_at',
      'google.calendar.connected_email',
      'google.calendar.connected_hd',
      'google.calendar.connected_at',
      'google.calendar.last_error',
      'google.calendar.last_success_at',
    ]);
}
