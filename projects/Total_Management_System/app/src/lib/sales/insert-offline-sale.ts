/* eslint-disable @typescript-eslint/no-explicit-any */
/**
 * offline_sales 판매 레코드 생성 — 판매번호(OS-YYYYMMDD-NNN) 채번 SSOT.
 *
 * 규칙: 같은 날짜 접두어(`OS-{날짜}-`) 중 **실제 최대 seq + 1**.
 *   - 과거 '당일 count+1' 방식은 날짜 소급편집/취소로 count가 되감기면 기존 번호와 충돌(23505)했음 → 폐기.
 * 안전장치: INSERT가 sale_number 중복(Postgres 23505)이면 다음 번호로 자동 재시도(동시등록·gap 방어).
 *
 * 3경로 공통 사용: 수동 판매등록(api/sales) · 이벤트 자동전환(convert-to-sale) · (레거시)이카운트 임포트.
 *
 * @param db       Supabase 클라이언트(서비스/서버)
 * @param saleDate 판매일 'YYYY-MM-DD'
 * @param payload  offline_sales insert 값 (sale_number 제외 — 여기서 채번해 붙임)
 * @returns 생성된 판매 레코드(select().single() 결과)
 */
export async function insertOfflineSale(
  db: any,
  saleDate: string,
  payload: Record<string, unknown>,
): Promise<any> {
  const day = String(saleDate || '').replace(/-/g, '').slice(0, 8);
  const prefix = `OS-${day}-`;

  // 접두어 최대 seq 조회
  const { data: rows } = await db
    .from('offline_sales')
    .select('sale_number')
    .like('sale_number', `${prefix}%`);
  let max = 0;
  for (const r of (rows || [])) {
    const m = /-(\d+)$/.exec(String(r?.sale_number ?? ''));
    if (m) { const n = parseInt(m[1], 10); if (!isNaN(n) && n > max) max = n; }
  }
  let seq = max + 1;

  // 중복 시 다음 번호로 재시도
  let created: any = null;
  let lastErr: any = null;
  for (let attempt = 0; attempt < 12; attempt++) {
    const saleNumber = `${prefix}${String(seq).padStart(3, '0')}`;
    const res = await db.from('offline_sales').insert({ sale_number: saleNumber, ...payload }).select().single();
    if (!res.error) { created = res.data; lastErr = null; break; }
    lastErr = res.error;
    if (res.error.code === '23505') { seq += 1; continue; } // 번호 중복 → 다음 번호
    break; // 그 외 오류는 즉시 중단
  }
  if (!created) throw lastErr || new Error('판매 등록 실패(sale_number 채번)');
  return created;
}
