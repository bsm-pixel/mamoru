/**
 * 판매·납품 기간 필터 계산 (SSOT, 2026-09-28)
 *
 * 왜 따로 뺐나: 목록은 기간을 적용하는데 **탭 배지는 전체 기간을 세고 있었다**.
 *   → '처리 필요 5' 인데 목록엔 4건(나머지 1건은 지난달 건). 숫자를 믿을 수 없게 된다.
 *   같은 계산을 두 곳이 나눠 갖지 않도록 여기 하나만 둔다.
 *
 * ⚠️ 로컬(KST) 달력 기준으로 계산한다. toISOString()(UTC)을 쓰면 날짜가 하루/한 달 밀린다
 *    (전에 '이번달'이 5월까지 나오던 버그의 원인).
 */

export type SalesDateRange = 'all' | 'today' | 'week' | 'month';

const pad = (n: number) => String(n).padStart(2, '0');
const ymd = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

/** 프리셋·커스텀 입력 → 실제 적용할 날짜 경계. 없으면 undefined(=제한 없음) */
export function resolveDateRange(filters?: {
  dateRange?: SalesDateRange | string;
  dateFrom?: string;
  dateTo?: string;
}): { from?: string; to?: string } {
  const to = filters?.dateTo || undefined;

  // 커스텀 시작일이 있으면 프리셋보다 우선 (목록 쿼리와 동일한 우선순위)
  if (filters?.dateFrom) return { from: filters.dateFrom, to };

  const range = filters?.dateRange;
  if (!range || range === 'all') return { to };

  const now = new Date();
  if (range === 'today') return { from: ymd(now), to };
  if (range === 'week') {
    const d = new Date(now);
    const dow = d.getDay();                              // 0=일
    d.setDate(d.getDate() + (dow === 0 ? -6 : 1 - dow)); // 이번주 월요일
    return { from: ymd(d), to };
  }
  return { from: `${now.getFullYear()}-${pad(now.getMonth() + 1)}-01`, to };  // 이번달 1일
}
