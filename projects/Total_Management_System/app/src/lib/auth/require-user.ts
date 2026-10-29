/**
 * 관리자 API 로그인 확인 (2026-10-29)
 *
 * 미들웨어 matcher 가 /api 를 제외하므로(src/middleware.ts) API 는 각 라우트가 스스로 로그인을 확인해야 한다.
 * 서비스 권한(createServiceClient)으로 DB 를 읽는 라우트가 이 확인을 빠뜨려 로그인 없이 매출·재고·이벤트가
 * 열려 있던 것을 막는다. 사용: `const deny = await requireUser(); if (deny) return deny;`
 * 공개 API(고객 폼·크론·웹훅)는 쓰지 않는다 — 각자 토큰/시크릿으로 확인.
 */
import { NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';

export async function requireUser(): Promise<NextResponse | null> {
  try {
    const supabase = await createServerSupabaseClient();
    const { data: { user } } = await supabase.auth.getUser();
    return user ? null : NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  } catch {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }
}
