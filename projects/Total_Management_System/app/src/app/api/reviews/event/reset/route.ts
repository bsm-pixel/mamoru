import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient, createServiceClient } from '@/lib/supabase/server';
import { inspectReset, performReset } from '@/lib/reviews/event-reset';

/**
 * 리뷰 이벤트 추첨 초기화 (2026-10-02) — 설정 > 시스템 「리뷰 추첨 초기화」
 *
 * 왜 설정 화면에 두나: 추첨 화면은 녹화·화면공유 대상이라 "다시 뽑기"류 버튼이 보이면
 * 조작으로 오해받는다(사장님 결정). 초기화는 추첨 화면 밖에서만 한다.
 *
 *   GET  ?month=YYMM                 → 미리보기 (당첨 인원·진행 상황·초기화 가능 여부)
 *   POST { month, confirm: 'YYMM' }  → 당첨 표시 해제 + 당첨자 배송 행 정리(soft)
 *
 * 🔒 막는 경우 (되돌리면 고객이 혼란 — 이미 밖으로 나간 것):
 *   - 이미 발표(announced)된 달
 *   - 당첨 안내를 보냈거나 / 고객이 주소를 저장했거나 / 송장이 있는 당첨자가 1명이라도 있음
 * 상품·마감일·응모 시작일 설정은 건드리지 않는다. 이력은 system_settings `review_event.reset_log` 에 남긴다.
 */

async function authUser() {
  const supabase = await createServerSupabaseClient();
  const { data: { user } } = await supabase.auth.getUser();
  return user;
}

export async function GET(req: NextRequest) {
  try {
    if (!(await authUser())) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const month = req.nextUrl.searchParams.get('month') || '';
    if (!/^\d{4}$/.test(month)) return NextResponse.json({ error: 'month(YYMM) required' }, { status: 400 });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    return NextResponse.json(await inspectReset(db, month));
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}

export async function POST(req: NextRequest) {
  try {
    const user = await authUser();
    if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
    const body = await req.json().catch(() => ({}));
    const month = String(body.month || '');
    if (!/^\d{4}$/.test(month)) return NextResponse.json({ error: 'month(YYMM) required' }, { status: 400 });
    // 실수 방지: 화면에서 고른 달을 한 번 더 보내야 한다
    if (String(body.confirm || '') !== month) return NextResponse.json({ error: '확인 값이 일치하지 않습니다' }, { status: 400 });

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;
    const result = await performReset(db, month, user.email || user.id);   // 🔒 서버에서 다시 판정 (화면 값 신뢰 X)
    if (result.blocked) return NextResponse.json({ error: result.blocked }, { status: 409 });
    return NextResponse.json(result);
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
