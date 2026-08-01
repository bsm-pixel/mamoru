/**
 * POST /api/push/test
 * 푸시 알림 발송 테스트 (관리자 전용)
 *
 * Body: { type: 'generic' | 'review' | 'consultation' | 'field_request' | 'talk_received'
 *                | 'field_confirmed' | 'field_reschedule' | 'repair_received' | 'order_received' }
 *
 * 동작:
 *   - 실제 sendPushToAll을 호출 → 관리자 기기로 진짜 푸시 전송
 *   - 제목에 "[테스트]" 접두어 붙여 구분
 *   - 운영과 동일하게 무조건 발송(2026-08-01 게이팅 제거) — settingKey 는 무시됨
 */

import { NextRequest, NextResponse } from 'next/server';
import { createServerSupabaseClient } from '@/lib/supabase/server';
import { sendPushToAll } from '@/lib/firebase/send-push';

type TestType =
  | 'generic'
  | 'review'
  | 'consultation'
  | 'field_request'
  | 'talk_received'
  | 'field_confirmed'
  | 'field_reschedule'
  | 'repair_received'
  | 'order_received';

const TEST_PRESETS: Record<TestType, { title: string; body: string; url: string; tag: string; settingKey?: string }> = {
  generic: {
    title: '[테스트] MAMORU TMS 알림',
    body: '푸시 알림이 정상 작동하는지 확인 중입니다.',
    url: '/dashboard',
    tag: 'mamoru-test',
    // settingKey 없음 — 토글 무관하게 무조건 발송 (기본 확인용)
  },
  review: {
    title: '[테스트] 새 리뷰 도착 ⭐5',
    body: '홍길동님 상담 리뷰 — 정말 친절하고 만족스러운 상담이었습니다...',
    url: '/reviews',
    tag: 'mamoru-test-review',
    settingKey: 'push.review_submitted',
  },
  consultation: {
    title: '[테스트] 새 상담 접수',
    body: '홍길동님 매장방문 상담 접수',
    url: '/consultations',
    tag: 'mamoru-test-consultation',
    settingKey: 'push.consultation_received',
  },
  field_request: {
    title: '[테스트] 새 출장 상담 접수',
    body: '김철수님 출장 상담 접수 (강남구)',
    url: '/consultations',
    tag: 'mamoru-test-field',
    settingKey: 'push.field_request',
  },
  talk_received: {
    title: '[테스트] 새 톡상담 접수',
    body: '박민수님 톡상담 접수',
    url: '/consultations',
    tag: 'mamoru-test-talk',
    settingKey: 'push.talk_received',
  },
  field_confirmed: {
    title: '[테스트] 출장 일정 확정 ✅',
    body: '김철수님이 2026-04-25 14:00로 확정했습니다',
    url: '/consultations',
    tag: 'mamoru-test-field-confirmed',
    settingKey: 'push.field_confirmed',
  },
  field_reschedule: {
    title: '[테스트] 출장 일정 재요청 🔄',
    body: '김철수님 재요청 — 오전에 다른 일정이 생겼어요',
    url: '/consultations',
    tag: 'mamoru-test-field-resched',
    settingKey: 'push.field_reschedule',
  },
  repair_received: {
    title: '[테스트] 새 복원수리 접수',
    body: '박민수님 복원수리 접수',
    url: '/repairs',
    tag: 'mamoru-test-repair',
    settingKey: 'push.repair_received',
  },
  order_received: {
    title: '[테스트] 아임웹 주문 접수 📦',
    body: '테스트 주문 1건 (58,000원)',
    url: '/orders',
    tag: 'mamoru-test-order',
    settingKey: 'push.order_received',
  },
};

export async function POST(req: NextRequest) {
  const supabase = await createServerSupabaseClient();
  const {
    data: { user },
  } = await supabase.auth.getUser();
  if (!user) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  try {
    const body = await req.json().catch(() => ({}));
    const type = (body.type as TestType) || 'generic';
    const preset = TEST_PRESETS[type];

    if (!preset) {
      return NextResponse.json(
        { ok: false, error: `알 수 없는 타입: ${type}` },
        { status: 400 },
      );
    }

    // 🔊 tag를 매 발송 유니크하게 — 같은 tag면 안드로이드가 기존 알림을 '무음 교체'해서
    //    두 번째부터 소리가 안 남(requireInteraction 로 알림이 안 사라져 계속 겹침). 테스트는 매번 새 알림 = 매번 소리.
    const result = await sendPushToAll({ ...preset, tag: `${preset.tag}-${Date.now()}` });

    return NextResponse.json({
      ok: true,
      data: {
        type,
        preset: { title: preset.title, settingKey: preset.settingKey || '(none)' },
        sent: result.sent,
        failed: result.failed,
      },
    });
  } catch (err) {
    return NextResponse.json({ ok: false, error: String(err) }, { status: 500 });
  }
}
