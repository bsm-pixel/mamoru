/* eslint-disable @typescript-eslint/no-explicit-any */
import { vi, beforeEach } from 'vitest';
import { FakeDb } from './fake-db';

/**
 * 자동 점검 공통 준비 — 운영에 닿는 길을 전부 막는다.
 *  ① 수파베이스 접속 → 메모리 가짜 저장소
 *  ② 알림톡·푸시 → 보내지 않고 "보내려 했다"만 기록
 *  ③ 그 밖의 모든 외부 호출(fetch) → 차단하고 기록
 */
export const ctx = {
  db: new FakeDb(),
  sent: [] as any[],        // 보내려던 알림톡
  pushes: [] as any[],      // 보내려던 관리자 푸시
  blocked: [] as string[],  // 차단된 외부 호출 주소
  afters: [] as Promise<unknown>[],
};

process.env.NEXT_PUBLIC_SUPABASE_URL ||= 'https://fake.local';
process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY ||= 'fake';
process.env.SUPABASE_SERVICE_ROLE_KEY ||= 'fake';

vi.mock('@/lib/supabase/server', () => ({
  createServerSupabaseClient: async () => ctx.db,
  createServiceClient: () => ctx.db,
}));

vi.mock('next/server', async (orig) => {
  const m = await orig<any>();
  return { ...m, after: (fn: () => unknown) => { ctx.afters.push(Promise.resolve().then(fn).catch(() => {})); } };
});

vi.mock('@/lib/notification/make-webhook', async (orig) => {
  const m = await orig<any>();
  return { ...m, sendNotification: vi.fn(async (p: any) => { ctx.sent.push(p); return { success: true }; }) };
});

vi.mock('@/lib/firebase/send-push', () => ({
  sendPushToAll: vi.fn(async (p: any) => { ctx.pushes.push(p); return { sent: 0 }; }),
}));

beforeEach(() => {
  ctx.db = new FakeDb();
  ctx.sent.length = 0; ctx.pushes.length = 0; ctx.blocked.length = 0; ctx.afters.length = 0;
  vi.stubGlobal('fetch', async (url: any) => {
    ctx.blocked.push(String(url));
    throw new Error('자동 점검: 외부 호출 차단 — ' + String(url));
  });
});

/** 응답 뒤로 미뤄 둔 작업(after)까지 끝내기 */
export async function settle() { await Promise.all(ctx.afters); await new Promise((r) => setTimeout(r, 0)); }

/** API 경로 직접 호출용 요청 만들기 */
export function jsonReq(url: string, method: string, body?: unknown) {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { NextRequest } = require('next/server');
  return new NextRequest('http://localhost' + url, {
    method,
    headers: { 'content-type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
}
export const idParam = (id: string) => ({ params: Promise.resolve({ id }) });
