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
          cookiesToSet.forEach(({ name, value }) =>
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

  // 로그인 페이지에 인증된 사용자 → 대시보드로 리다이렉트
  if (user && pathname.startsWith('/login')) {
    const url = request.nextUrl.clone();
    url.pathname = '/dashboard';
    return NextResponse.redirect(url);
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
    /*
     * 인증이 필요한 경로만 매칭 — 공개 API는 matcher에서 제외하여 미들웨어 자체가 실행 안 됨
     * 제외 목록:
     * - _next/static, _next/image, favicon 등 정적 파일
     * - /api/consultation/public/* (고객 접수/슬롯/설정)
     * - /api/repair/public/* (고객 접수/공휴일)
     * - /api/repair/report (공개 리포트)
     * - /api/repair/sync (GAS 동기화)
     * - /api/consultation/sync, /api/consultation/notify
     * - /api/reviews/* (공개 리뷰 API)
     * - /api/cron/* (Vercel Cron)
     * - /api/imweb/* (아임웹 연동)
     * - /api/verify/* (QR 인증)
     * - /login, /contract, /diagnosis
     */
    '/((?!_next/static|_next/image|favicon.ico|manifest.*\\.json|icon-.*\\.png|sw\\.js|api/consultation/public|api/repair/public|api/repair/report|api/repair/sync|api/consultation/sync|api/consultation/notify|api/reviews|api/cron|api/imweb|api/verify|login|contract|diagnosis).*)',
  ],
};
