import { createServerClient } from '@supabase/ssr';
import { NextResponse, type NextRequest } from 'next/server';

export async function middleware(request: NextRequest) {
  let supabaseResponse = NextResponse.next({ request });

  const supabase = createServerClient(
    process.env.NEXT_PUBLIC_SUPABASE_URL!,
    process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!,
    {
      cookies: {
        getAll() {
          return request.cookies.getAll();
        },
        setAll(cookiesToSet) {
          cookiesToSet.forEach(({ name, value, options }) =>
            request.cookies.set(name, value)
          );
          supabaseResponse = NextResponse.next({ request });
          cookiesToSet.forEach(({ name, value, options }) =>
            supabaseResponse.cookies.set(name, value, options)
          );
        },
      },
    }
  );

  const {
    data: { user },
  } = await supabase.auth.getUser();

  const { pathname } = request.nextUrl;

  // 인증 불필요 경로 — 고객 공개 API + 동기화 + Cron
  // consultation/public, repair/public 추가 (2026-03-31)
  if (pathname.startsWith('/login') || pathname.startsWith('/contract') || pathname.startsWith('/diagnosis') || pathname.startsWith('/api/cron') || pathname.startsWith('/api/consultation/sync') || pathname.startsWith('/api/consultation/notify') || pathname.startsWith('/api/consultation/public') || pathname.startsWith('/api/repair/report') || pathname.startsWith('/api/repair/sync') || pathname.startsWith('/api/repair/public') || pathname.startsWith('/api/reviews/submit') || pathname.startsWith('/api/reviews/info') || pathname.startsWith('/api/reviews/upload') || pathname.startsWith('/api/reviews/public')) {
    if (user && pathname.startsWith('/login')) {
      const url = request.nextUrl.clone();
      url.pathname = '/dashboard';
      return NextResponse.redirect(url);
    }
    return supabaseResponse;
  }

  // 미인증 → 로그인으로 리다이렉트
  if (!user) {
    const url = request.nextUrl.clone();
    url.pathname = '/login';
    url.searchParams.set('redirect', pathname);
    return NextResponse.redirect(url);
  }

  return supabaseResponse;
}

export const config = {
  matcher: [
    '/((?!_next/static|_next/image|favicon.ico|manifest.*\\.json|icon-.*\\.png|sw\\.js).*)',
  ],
};
