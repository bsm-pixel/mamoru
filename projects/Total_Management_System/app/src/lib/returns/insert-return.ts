/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * returns 반품·교환수거 레코드 생성 — 반품번호(RT-YYYYMMDD-NNN) 채번 SSOT.
 * insert-offline-sale.ts 와 동일한 "접두어 최대 seq + 1 + 23505 재시도" 방식(동시등록·gap 방어).
 *
 * @param db      Supabase 클라이언트(서버)
 * @param dayISO  기준일 'YYYY-MM-DD' (보통 오늘)
 * @param payload returns insert 값 (return_number 제외 — 여기서 채번)
 */
export async function insertReturn(
  db: any,
  dayISO: string,
  payload: Record<string, unknown>,
): Promise<any> {
  const day = String(dayISO || '').replace(/-/g, '').slice(0, 8);
  const prefix = `RT-${day}-`;

  const { data: rows } = await db
    .from('returns')
    .select('return_number')
    .like('return_number', `${prefix}%`);
  let max = 0;
  for (const r of (rows || [])) {
    const m = /-(\d+)$/.exec(String(r?.return_number ?? ''));
    if (m) { const n = parseInt(m[1], 10); if (!isNaN(n) && n > max) max = n; }
  }
  let seq = max + 1;

  let created: any = null;
  let lastErr: any = null;
  for (let attempt = 0; attempt < 12; attempt++) {
    const returnNumber = `${prefix}${String(seq).padStart(3, '0')}`;
    const res = await db.from('returns').insert({ return_number: returnNumber, ...payload }).select().single();
    if (!res.error) { created = res.data; lastErr = null; break; }
    lastErr = res.error;
    if (res.error.code === '23505') { seq += 1; continue; }
    break;
  }
  if (!created) throw lastErr || new Error('반품 등록 실패(return_number 채번)');
  return created;
}
