import { createServerClient } from '@supabase/ssr';
import { NextResponse, type NextRequest } from 'next/server';

/** 인증 불필요 공개 경로 */
const PUBLIC_PATHS = [
  '/login',
  '/contract',
  '/diagnosis',
  '/api/cron',
  '/api/consultation/sync',
  '/api/consultation/notify',
  '/api/consultation/public',
  '/api/repair/report',
  '/api/repair/sync',
  '/api/repair/public',
  '/api/reviews/submit',
  '/api/reviews/info',
  '/api/reviews/upload',
  '/api/reviews/public',
  '/api/imweb/orders',
];

function isPublicPath(pathname: string): boolean {
  return PUBLIC_PATHS.some(p => pathname.startsWith(p));
}

export async function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;

  // 공개 경로 → 인증 없이 통과 (getUser 호출 불필요)
  if (isPublicPath(pathname)) {
    return NextResponse.next({ request });
  }

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
    '/((?!_next/static|_next/image|favicon.ico|manifest.*\\.json|icon-.*\\.png|sw\\.js).*)',
  ],
};
