/**
 * 리뷰 이벤트 추첨 초기화 — 판정·실행 로직 (2026-10-02)
 * 라우트(api/reviews/event/reset)가 인증 후 호출한다. 판정은 항상 서버에서 다시 한다.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export async function inspectReset(db: Db, month: string) {
  const [{ data: cfg }, { data: winners }, { data: ships }] = await Promise.all([
    db.from('review_event_config').select('status').eq('month', month).maybeSingle(),
    db.from('reviews').select('id, event_rank').eq('event_month', month).not('event_rank', 'is', null),
    db.from('review_event_shipments')
      .select('id, won_notified_at, address_submitted_at, invoice_number, cancelled_at')
      .eq('event_month', month).is('cancelled_at', null),
  ]);
  const byRank: Record<string, number> = {};
  (winners || []).forEach((w: { event_rank: number }) => { byRank[w.event_rank] = (byRank[w.event_rank] || 0) + 1; });
  const s = ships || [];
  const notified = s.filter((x: { won_notified_at: string | null }) => x.won_notified_at).length;
  const addressed = s.filter((x: { address_submitted_at: string | null }) => x.address_submitted_at).length;
  const invoiced = s.filter((x: { invoice_number: string | null }) => x.invoice_number).length;
  const status: string = cfg?.status || 'none';

  let blocked: string | null = null;
  if (status === 'announced') blocked = '이미 고객 페이지에 발표된 달입니다. 초기화하려면 리뷰 이벤트 관리에서 먼저 [비공개]로 바꾸세요.';
  else if (invoiced > 0) blocked = `송장이 생성된 당첨자가 ${invoiced}명 있습니다. 당첨자 배송에서 송장을 먼저 취소하세요.`;
  else if (addressed > 0) blocked = `배송지를 저장한 당첨자가 ${addressed}명 있습니다. 이미 고객이 참여한 추첨은 초기화할 수 없습니다.`;
  else if (notified > 0) blocked = `당첨 안내를 보낸 당첨자가 ${notified}명 있습니다. 이미 안내한 추첨은 초기화할 수 없습니다.`;

  return { month, status, winners: (winners || []).length, byRank, shipments: s.length, notified, addressed, invoiced, blocked };
}

/** 초기화 실행. 막힌 달이면 { blocked } 만 돌려주고 아무것도 바꾸지 않는다 */
export async function performReset(db: Db, month: string, by: string): Promise<{ ok: boolean; blocked?: string; cleared: number; remaining?: number; note?: string }> {
  const before = await inspectReset(db, month);          // 🔒 서버에서 다시 판정 (화면 값 신뢰 X)
  if (before.blocked) return { ok: false, blocked: before.blocked, cleared: 0 };
  if (before.winners === 0 && before.shipments === 0) return { ok: true, cleared: 0, note: '초기화할 당첨자가 없습니다' };

  const now = new Date().toISOString();
  // 1) 당첨 표시 해제 (저장 API 와 같은 4개 필드)
  const { error: e1 } = await db.from('reviews')
    .update({ event_month: null, event_rank: null, event_display_name: null, event_route: null })
    .eq('event_month', month);
  if (e1) throw e1;
  // 2) 당첨자 배송 행 정리 — 삭제 대신 취소 표시(soft). 아무것도 발송 안 된 행만 남아 있다(위에서 확인)
  const { error: e2 } = await db.from('review_event_shipments')
    .update({ cancelled_at: now, updated_at: now, memo: `추첨 초기화 (${now.slice(0, 10)})` })
    .eq('event_month', month).is('cancelled_at', null);
  if (e2 && !/does not exist|schema cache/i.test(e2.message || '')) throw e2;

  // 3) 이력 — 누가·언제·몇 명 (최근 20건)
  try {
    const { data: logRow } = await db.from('system_settings').select('value').eq('key', 'review_event.reset_log').maybeSingle();
    const prev = Array.isArray(logRow?.value) ? logRow.value : [];
    const next = [{ month, at: now, by, winners: before.winners, byRank: before.byRank }, ...prev].slice(0, 20);
    await db.from('system_settings').upsert({ key: 'review_event.reset_log', value: next, updated_at: now }, { onConflict: 'key' });
  } catch { /* 이력 실패가 초기화를 막지 않게 */ }

  const after = await inspectReset(db, month);
  return { ok: true, cleared: before.winners, remaining: after.winners };
}
