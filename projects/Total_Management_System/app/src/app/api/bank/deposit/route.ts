import { NextRequest, NextResponse, after } from 'next/server';
import { createServiceClient } from '@/lib/supabase/server';
import { parseWooriDeposit, normalizeName } from '@/lib/bank/woori-deposit';
import { sendRepairPaymentConfirmed } from '@/lib/repair/payment-notify';
import { sendPushToAll } from '@/lib/firebase/send-push';
import { formatKRW } from '@/lib/utils/format';

/**
 * POST /api/bank/deposit — 은행 입금 알림 → 복원수리 입금확인 자동화 (161, 2026-10-13)
 *
 * 사장님 안드로이드 폰의 MacroDroid 가 「우리WON뱅킹」 알림을 받으면 이 주소로 원문을 보낸다.
 *   인증: 헤더 x-mamoru-key 또는 ?key= == BANK_WEBHOOK_SECRET (로그인 세션 없음)
 *   본문: JSON {"title": "...", "text": "..."} 또는 알림 글자 그대로(text/plain)
 *
 * 처리 규칙 — 확실할 때만 고객에게 알림톡이 나간다:
 *   ① 입금대기 수리 건(paid_at 없음·취소 아님) 중 「이름+총액」 정확히 1건 → 입금 처리 + 입금확인 알림톡
 *   ② 총액만 맞는 건이 있거나 이름+총액이 2건 이상 → 알림톡 X, 사장님 푸시 「입금 확인 필요」
 *   ③ 맞는 건 없음(제품 판매·개인 송금 등) → 기록만 (푸시 X — 알림 피로 방지)
 * 받은 알림은 모두 bank_deposits 에 남고, 같은 알림이 두 번 와도 한 번만 처리된다(dedupe_key).
 */
export async function POST(req: NextRequest) {
  const secret = process.env.BANK_WEBHOOK_SECRET;
  const key = req.headers.get('x-mamoru-key') || req.nextUrl.searchParams.get('key');
  if (!secret || key !== secret) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  try {
    const raw = await req.text();
    let text = raw;
    try {
      const j = JSON.parse(raw);
      text = [j.title, j.text].filter(Boolean).join('\n');
    } catch { /* JSON 아니면 원문 그대로 */ }

    const parsed = parseWooriDeposit(text);
    if (!parsed) return NextResponse.json({ outcome: 'ignored' });  // 출금·광고 등

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const db = createServiceClient() as any;

    // 같은 알림 재전송 → 무시 (unique dedupe_key)
    const ins = await db.from('bank_deposits').insert({
      bank: 'woori',
      depositor: parsed.depositor,
      amount: parsed.amount,
      deposited_at: parsed.depositedAt,
      raw_text: text.slice(0, 1000),
      dedupe_key: parsed.dedupeKey,
      outcome: 'pending',
    }).select('id').single();
    if (ins.error) {
      if (ins.error.code === '23505') return NextResponse.json({ outcome: 'duplicate' });
      throw ins.error;
    }
    const depositId: string = ins.data.id;

    // 입금대기 + 총액 일치 후보
    const { data: cands, error: cErr } = await db
      .from('repairs')
      .select('id, as_id, name, phone, status, service_cost, shipping_fee, total_amount, invoice_number')
      .is('paid_at', null)
      .neq('status', 'cancelled')
      .eq('total_amount', parsed.amount);
    if (cErr) throw cErr;
    const amountMatches: Array<{ id: string; as_id: string; name: string; status: string } & Record<string, unknown>> = cands || [];
    const nameMatches = amountMatches.filter((r) => normalizeName(r.name) === normalizeName(parsed.depositor));
    const who = `${parsed.depositor} ${formatKRW(parsed.amount)}`;

    // ① 정확히 1건 → 자동 입금확인
    if (nameMatches.length === 1) {
      const target = nameMatches[0];
      // 사장님이 동시에 수동 처리했으면 건너뜀 (paid_at IS NULL 조건부 갱신)
      const { data: claimed } = await db
        .from('repairs')
        .update({ paid_at: new Date().toISOString(), payment_method: 'transfer' })
        .eq('id', target.id)
        .is('paid_at', null)
        .select('as_id, name, phone, service_cost, shipping_fee, total_amount, invoice_number')
        .maybeSingle();

      if (claimed) {
        await db.from('repair_history').insert({
          repair_id: target.id,
          from_status: target.status,
          to_status: target.status,
          changed_by: null,
          note: `자동 입금확인 (우리은행 알림): ${who}`,
        });
        await db.from('bank_deposits')
          .update({ outcome: 'matched', repair_id: target.id, resolved_at: new Date().toISOString() })
          .eq('id', depositId);
        after(async () => {
          await sendRepairPaymentConfirmed(claimed);
          await sendPushToAll({
            title: '✅ 입금 자동확인',
            body: `${claimed.name}님 ${formatKRW(parsed.amount)} · ${claimed.as_id} — 입금확인 알림톡 발송`,
            url: '/repairs',
            tag: `mamoru-bank-${depositId}`,
            fallback: false, // 정보성 — 메일 예비발송 불필요
          });
        });
        return NextResponse.json({ outcome: 'matched', as_id: claimed.as_id });
      }
      // 그 사이 사장님이 수동으로 입금처리함 → 알림톡은 수동 경로에서 이미 나감
      await db.from('bank_deposits')
        .update({ outcome: 'already_paid', repair_id: target.id, resolved_at: new Date().toISOString() })
        .eq('id', depositId);
      return NextResponse.json({ outcome: 'already_paid' });
    }

    // ② 애매함 → 사장님 확인
    if (amountMatches.length > 0) {
      const list = (nameMatches.length > 1 ? nameMatches : amountMatches)
        .slice(0, 3).map((r) => `${r.name}(${r.as_id})`).join(', ');
      await db.from('bank_deposits')
        .update({ outcome: 'needs_review', candidate_repair_ids: amountMatches.map((r) => r.id) })
        .eq('id', depositId);
      after(async () => {
        await sendPushToAll({
          title: '💰 입금 확인 필요',
          body: `${who} — 금액 맞는 복원수리: ${list}. 확인 후 입금처리해 주세요`,
          url: '/repairs',
          tag: `mamoru-bank-${depositId}`,
        });
      });
      return NextResponse.json({ outcome: 'needs_review', candidates: amountMatches.length });
    }

    // ③ 복원수리와 무관한 입금 → 기록만
    await db.from('bank_deposits').update({ outcome: 'unmatched' }).eq('id', depositId);
    return NextResponse.json({ outcome: 'unmatched' });
  } catch (err) {
    console.error('[bank/deposit] 실패:', err);
    return NextResponse.json({ error: String(err) }, { status: 500 });
  }
}
