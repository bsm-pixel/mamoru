import { NextRequest, NextResponse } from 'next/server';
import crypto from 'crypto';
import { createServiceClient } from '@/lib/supabase/server';
import { displayWinnerName, maskPhoneEvent } from '@/lib/reviews/mask';

/**
 * GET /api/reviews/event-preview?key=<비밀키>[&month=YYMM]   (2026-10-02)
 *
 * 인스타 홍보물 제작용(Claude 스킬 「리뷰이벤트 홍보물」) **비공개 읽기 전용** 주소.
 * 공개 API(event-public)는 진행중/발표된 것만 주므로 "고객 공개 전에 게시물을 먼저 만든다"가 불가능했다.
 * 이 주소는 비공개(draft) 설정과 아직 게시 전인 당첨자까지 돌려준다 — 단,
 *   · 키(system_settings `review_event.preview_key`)가 맞아야 열린다. 키 없음/불일치 = 404 (존재 자체를 숨김)
 *   · 당첨자는 event-public 과 같은 마스킹만 (백*민 님 / (3562)). 실명·전체 번호는 절대 내보내지 않는다
 *   · 읽기 전용. 어떤 것도 바꾸지 않는다
 *
 * month 없으면 최근 3개 달을 돌려준다(최신순). 키 교체 = system_settings 값만 바꾸면 즉시 반영.
 */

export const dynamic = 'force-dynamic';

interface Prize { rank: number; label?: string; name?: string; desc?: string; image_url?: string; image_urls?: string[]; count?: number }
interface Cfg { month: string; status: string; deadline: string | null; announce_at: string | null; entry_start: string | null; prizes: Prize[] | null }
interface Win { event_month: string; event_rank: number; name: string | null; phone: string | null; event_display_name: string | null; content: string | null; stars: number | null; created_at: string }

const NOT_FOUND = () => NextResponse.json({ error: 'not_found' }, { status: 404, headers: { 'Cache-Control': 'no-store' } });

function monthLabel(m: string): string {
  return `${parseInt(m.slice(2, 4), 10)}월`;
}

function safeEqual(a: string, b: string): boolean {
  const x = Buffer.from(a), y = Buffer.from(b);
  return x.length === y.length && crypto.timingSafeEqual(x, y);
}

export async function GET(req: NextRequest) {
  try {
    const key = (req.nextUrl.searchParams.get('key') || '').trim();
    const month = (req.nextUrl.searchParams.get('month') || '').trim();
    if (key.length < 16) return NOT_FOUND();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const { data: keyRow } = await db.from('system_settings').select('value').eq('key', 'review_event.preview_key').maybeSingle();
    const expected = keyRow?.value ? String(keyRow.value).replace(/^"|"$/g, '') : '';
    if (!expected || !safeEqual(key, expected)) return NOT_FOUND();

    let q = db.from('review_event_config')
      .select('month, status, deadline, announce_at, entry_start, prizes')
      .order('month', { ascending: false });
    q = /^\d{4}$/.test(month) ? q.eq('month', month) : q.limit(3);
    const { data: cfgs, error } = await q;
    if (error) throw error;
    const list: Cfg[] = cfgs || [];
    const months = list.map((c) => c.month);

    const winsByMonth: Record<string, Win[]> = {};
    if (months.length) {
      const { data: wins, error: wErr } = await db.from('reviews')
        .select('event_month, event_rank, name, phone, event_display_name, content, stars, created_at')
        .in('event_month', months).not('event_rank', 'is', null)
        .order('event_rank', { ascending: true }).order('created_at', { ascending: true });
      if (wErr) throw wErr;
      (wins as Win[] || []).forEach((w) => { (winsByMonth[w.event_month] ||= []).push(w); });
    }

    const events = list.map((c) => {
      const prizes = (Array.isArray(c.prizes) ? c.prizes : [])
        .filter((p) => (p.count || 0) > 0 && (p.name || '').trim())      // 인원 0·상품명 빈 등수 제외 (고객 화면과 같은 규칙)
        .map((p) => ({
          rank: p.rank,
          label: (p.label || '').trim() || `${p.rank}등`,
          count: p.count || 0,
          name: (p.name || '').trim(),
          desc: (p.desc || '').trim(),
          image_url: (Array.isArray(p.image_urls) && p.image_urls.filter(Boolean)[0]) || p.image_url || '',
        }));
      const labelOf = (rank: number) => ((Array.isArray(c.prizes) ? c.prizes : []).find((p) => p.rank === rank)?.label || '').trim() || `${rank}등`;
      const winners = (winsByMonth[c.month] || []).map((w) => ({
        rank: w.event_rank,
        rank_label: labelOf(w.event_rank),
        name: displayWinnerName(w.name, w.event_display_name),            // 🔒 마스킹된 값만
        phone: maskPhoneEvent(w.phone),                                    // 🔒 뒷 4자리만
        display: `${displayWinnerName(w.name, w.event_display_name)}${maskPhoneEvent(w.phone) ? ' ' + maskPhoneEvent(w.phone) : ''}`,
        review: (w.content || '').slice(0, 300),
        stars: w.stars || 0,
      }));
      return {
        month: c.month,
        label: monthLabel(c.month),                    // 예: "9월"
        status: c.status,                               // draft(비공개) | live(진행중) | announced(발표됨)
        published_to_customers: c.status !== 'draft',
        deadline: c.deadline,
        announce_at: c.announce_at,
        prizes,
        winners,
        winner_count: winners.length,
        winners_announced: c.status === 'announced',
      };
    });

    return NextResponse.json({ generated_at: new Date().toISOString(), events }, { headers: { 'Cache-Control': 'no-store' } });
  } catch (err) {
    console.error('[reviews/event-preview] 조회 실패:', err);
    return NextResponse.json({ error: 'server_error' }, { status: 500, headers: { 'Cache-Control': 'no-store' } });
  }
}
