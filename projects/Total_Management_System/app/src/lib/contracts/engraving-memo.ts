/**
 * 계약서 각인 문구 → 전환된 판매 메모에 반영 (2026-11-26, 작업실 13번)
 *
 * 계약서 → [판매 전환] 때 판매 메모에 「계약서 CT-… · 각인 「문구」」 또는 「각인 문구 대기」가 들어간다.
 * 전환 뒤에 고객이 링크로 문구를 보내거나 직원이 문구를 고치면, 전엔 판매 메모가 옛 값 그대로라
 * 출고 준비표 · 판매 상세에 「각인 문구 대기」가 계속 찍혔다. 문구가 바뀔 때마다 이 함수로 판매 메모를 맞춘다.
 */

/** 각인 상태 — on=false 면 각인 안 함 */
export interface EngravingState { on: boolean; text: string | null }

const ENG_RE = /각인 「[^」]*」|각인 문구 대기/;

/** 메모 안의 각인 부분만 새 상태로 바꾼다. 다른 글은 그대로 */
export function applyEngravingToMemo(memo: string | null | undefined, st: EngravingState): string | null {
  const cur = memo || '';
  const part = st.on ? (st.text ? `각인 「${st.text}」` : '각인 문구 대기') : '';
  if (ENG_RE.test(cur)) {
    if (part) return cur.replace(ENG_RE, part);
    // 각인을 끈 경우 — 그 부분과 앞뒤 구분자(·)만 뺀다
    const removed = cur.replace(/\s*·\s*(각인 「[^」]*」|각인 문구 대기)/, '').replace(/(각인 「[^」]*」|각인 문구 대기)\s*·\s*/, '').replace(ENG_RE, '').trim();
    return removed || null;
  }
  if (!part) return memo ?? null;
  return cur.trim() ? `${cur.trim()} · ${part}` : part;
}

/**
 * 계약서에 이어진 판매가 있으면 그 메모의 각인 부분을 맞춘다. 판매가 없으면 아무것도 안 한다.
 * 실패해도 계약서 저장은 그대로(로그만) — 부르는 쪽이 이미 저장을 끝낸 뒤 부른다.
 */
export async function syncSaleEngravingMemo(
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  db: any,
  contractId: string,
  st: EngravingState,
): Promise<void> {
  try {
    const { data: c } = await db.from('contracts').select('offline_sale_id').eq('id', contractId).maybeSingle();
    const saleId = c?.offline_sale_id as string | null | undefined;
    if (!saleId) return;
    const { data: s } = await db.from('offline_sales').select('memo').eq('id', saleId).maybeSingle();
    if (!s) return;
    const next = applyEngravingToMemo(s.memo, st);
    if (next === (s.memo ?? null)) return;
    const { error } = await db.from('offline_sales').update({ memo: next }).eq('id', saleId);
    if (error) console.error('[contract] 판매 메모 각인 반영 실패:', error.message);
  } catch (e) {
    console.error('[contract] 판매 메모 각인 반영 실패', e);
  }
}
