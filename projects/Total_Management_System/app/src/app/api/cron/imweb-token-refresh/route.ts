import { NextRequest, NextResponse } from 'next/server';
import { forceRefreshOpenApiToken } from '@/lib/imweb/client';
import { sendPushToAll } from '@/lib/firebase/send-push';

/**
 * GET /api/cron/imweb-token-refresh — 아임웹 OpenAPI 토큰 자동 유지 (Vercel Cron)
 *
 * 🅰 판매가 뜸해도 12시간마다 토큰을 미리 갱신 → 방치로 인한 재고 push 사망 방지.
 * 🅱 갱신 실패(재인증 필요) 시 사장님 앱에 즉시 푸시 알림 → 몰래 쌓이지 않게.
 */
export async function GET(request: NextRequest) {
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  const result = await forceRefreshOpenApiToken();

  if (!result.ok) {
    // 🅱 끊김 → 즉시 알림 (설정 안 들어가도 사장님이 바로 인지)
    try {
      await sendPushToAll({
        title: '⚠️ 아임웹 재고 연동 끊김',
        body: '재고가 아임웹에 자동 반영되지 않습니다. 설정 > 상품·재고에서 [아임웹 재연결]을 눌러 복구하세요.',
        url: '/settings',
        tag: 'imweb-sync-broken',
      });
    } catch (e) {
      console.error('[cron/imweb-token] 알림 발송 실패:', e);
    }
    console.error('[cron/imweb-token] 토큰 갱신 실패:', result.error);
    return NextResponse.json({ ok: false, alerted: true, error: result.error });
  }

  return NextResponse.json({ ok: true, updatedAt: result.updatedAt });
}
