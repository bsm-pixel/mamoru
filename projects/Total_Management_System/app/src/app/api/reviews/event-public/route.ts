import { NextResponse } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { displayWinnerName, maskPhoneEvent } from '@/lib/reviews/mask';

/**
 * GET /api/reviews/event-public
 * 리뷰 이벤트 고객 페이지(page_review_event.html)가 fetch 하는 공개 읽기전용 엔드포인트.
 * - current : status='live' 인 이달 이벤트(상품/마감일/히어로) → Hero + 이달의 상품
 * - past    : status='announced' 인 지난 달들의 마스킹 당첨자 → 지난 당첨자 아카이브(월 탭)
 * PII 는 반드시 마스킹된 값만 반환. 발표(announced)된 것만 노출 → 전체 후기 유출 차단.
 */

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
  'Cache-Control': 'public, max-age=300, s-maxage=300',
};

export function OPTIONS() {
  return new NextResponse(null, { status: 204, headers: CORS_HEADERS });
}

interface Prize { rank: number; name?: string; desc?: string; image_url?: string; count?: number }
interface WinnerRow {
  event_month: string; event_rank: number;
  name: string | null; phone: string | null;
  event_display_name: string | null; event_route: string | null;
  content: string | null; photo_urls: string[] | null; stars: number | null;
  product: string | null; subtype: string | null; created_at: string;
}

/** 'YYMM' → 'N월' */
function monthLabel(m: string): string {
  const mm = parseInt(m.slice(2, 4), 10);
  return Number.isFinite(mm) ? `${mm}월` : m;
}

export async function GET() {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    // 1) 공개 대상 설정: live(진행중) + announced(발표됨)
    const { data: configs, error: cfgErr } = await db
      .from('review_event_config')
      .select('month, deadline, announce_at, hero_image_url, prizes, status')
      .in('status', ['live', 'announced'])
      .order('month', { ascending: false });
    if (cfgErr) throw cfgErr;

    const cfgList: Array<{ month: string; deadline: string | null; announce_at: string | null; hero_image_url: string | null; prizes: Prize[] | null; status: string }> = configs || [];

    // 2) current = live (최신 우선), past 후보 = announced
    const liveCfg = cfgList.find((c) => c.status === 'live') || null;
    const announcedMonths = cfgList.filter((c) => c.status === 'announced').map((c) => c.month);

    // 3) 발표된 달들의 당첨자(event_rank 있는 것만) 조회
    const winnersByMonth: Record<string, WinnerRow[]> = {};
    if (announcedMonths.length > 0) {
      const { data: wins, error: wErr } = await db
        .from('reviews')
        .select('event_month, event_rank, name, phone, event_display_name, event_route, content, photo_urls, stars, product, subtype, created_at')
        .in('event_month', announcedMonths)
        .not('event_rank', 'is', null)
        .order('event_rank', { ascending: true })
        .order('created_at', { ascending: true });
      if (wErr) throw wErr;
      (wins as WinnerRow[] | null || []).forEach((w) => {
        (winnersByMonth[w.event_month] ||= []).push(w);
      });
    }

    // 4) past 조립 (당첨자 있는 발표 달만, 월 내림차순)
    const past = cfgList
      .filter((c) => c.status === 'announced' && (winnersByMonth[c.month]?.length || 0) > 0)
      .map((c) => ({
        month: c.month,
        label: monthLabel(c.month),
        winners: winnersByMonth[c.month].map((w) => ({
          rank: w.event_rank,
          name: displayWinnerName(w.name, w.event_display_name),
          phone: maskPhoneEvent(w.phone),
          route: (w.event_route || w.product || '').trim(),
          review: w.content || '',
          photos: Array.isArray(w.photo_urls) ? w.photo_urls : [],
          stars: w.stars || 0,
        })),
      }));

    // 5) current 조립
    const current = liveCfg
      ? {
          month: liveCfg.month,
          label: monthLabel(liveCfg.month),
          deadline: liveCfg.deadline,
          announce_at: liveCfg.announce_at,
          hero_image_url: liveCfg.hero_image_url,
          prizes: Array.isArray(liveCfg.prizes) ? liveCfg.prizes : [],
        }
      : null;

    return NextResponse.json({ current, past }, { headers: CORS_HEADERS });
  } catch (err) {
    console.error('[reviews/event-public] 조회 실패:', err);
    return NextResponse.json({ current: null, past: [], error: String(err) }, { status: 500, headers: CORS_HEADERS });
  }
}
