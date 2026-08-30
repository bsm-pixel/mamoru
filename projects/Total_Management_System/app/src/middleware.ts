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
     * 🔒 캐치올: 아래를 제외한 '모든 페이지'를 인증 뒤에 둔다 (2026-08-30 하드닝)
     *   - 예전엔 경로를 하나씩 나열 → /sourcing·/schedule·/events·/manual-invoices·
     *     /deliveries·/stock-sale 가 목록에서 빠져 로그인 없이 껍데기가 열렸음.
     *   - 이제 새 페이지를 추가해도 자동으로 보호됨(재발 방지).
     * 제외(공개 유지):
     *   - api : 각 API가 자체 auth.getUser() 검사
     *   - _next/static·_next/image : 빌드 정적 자산
     *   - firebase-messaging-sw.js : FCM 서비스워커(무인증 로드 필수)
     *   - 정적 파일 확장자(png/svg/json/wav 등) : manifest·아이콘·알림음 등
     *   - /login 은 matcher엔 포함되나 미들웨어 로직이 예외 처리(무인증 접근 허용)
     */
    '/((?!api|_next/static|_next/image|favicon.ico|firebase-messaging-sw.js|.*\\.(?:png|jpg|jpeg|svg|gif|webp|ico|woff2?|json|wav|txt|xml)$).*)',
  ],
};
