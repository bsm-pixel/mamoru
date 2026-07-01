/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * 카테고리 접두어의 다음 SKU 계산 (SSOT).
 *
 * 주의: `like('{prefix}%')` 는 유사 접두어(예: prefix 'SU' → 'SUP001' 부자재)까지 잡으므로,
 *   반드시 정확히 `^{prefix}\d+$` 형식만 번호로 인정한다. (그러지 않으면 최대값 파싱이 NaN→1로 떨어져
 *   기존 SU001 과 duplicate key 충돌 — 2026-07-01 실사고)
 * 또한 혹시 모를 번호 충돌을 피하려 이미 사용 중인 번호는 건너뛴다.
 */
export async function computeNextSku(db: any, prefix: string): Promise<string> {
  const clean = String(prefix || '').trim();
  if (!clean) return '';

  const { data: rows } = await db
    .from('products')
    .select('sku')
    .like('sku', `${clean}%`);

  const escaped = clean.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const exact = new RegExp(`^${escaped}(\\d+)$`);

  const used = new Set<number>();
  let maxNum = 0;
  for (const r of (rows || [])) {
    const m = exact.exec(String(r?.sku ?? ''));
    if (m) {
      const n = parseInt(m[1], 10);
      if (!isNaN(n)) {
        used.add(n);
        if (n > maxNum) maxNum = n;
      }
    }
  }

  let nextNum = maxNum + 1;
  while (used.has(nextNum)) nextNum++;

  return `${clean}${String(nextNum).padStart(3, '0')}`;
}
