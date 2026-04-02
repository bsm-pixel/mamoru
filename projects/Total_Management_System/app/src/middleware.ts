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

  // 미인증 → 로그인으로 리다이렉트 (로그인 페이지 자체는 제외)
  if (!user && !pathname.startsWith('/login')) {
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
     * 대시보드 페이지만 매칭 — API 경로는 미들웨어 대상 아님
     * 인증이 필요한 TMS 페이지: /dashboard, /orders, /sales, /repairs 등
     */
    '/dashboard/:path*',
    '/orders/:path*',
    '/sales/:path*',
    '/repairs/:path*',
    '/consultations/:path*',
    '/products/:path*',
    '/customers/:path*',
    '/suppliers/:path*',
    '/supplies/:path*',
    '/purchasing/:path*',
    '/contracts/:path*',
    '/inventory/:path*',
    '/serials/:path*',
    '/reports/:path*',
    '/reviews/:path*',
    '/settings/:path*',
    '/login',
  ],
};
