/**
 * 외화 영수증 → 원화 추정 (2026-10-29)
 *
 * 카드사 청구 원화는 메일로 오지 않는다(4개월 메일함 확인: 카드사 승인·청구 메일 0통). 그래서 결제일 기준환율로 추정해
 * 확인 대기 금액 칸을 채우고, 사장님이 명세서 금액을 알면 [확인] 전에 고친다. 실제 청구액과 보통 1~2% 차이
 * (카드사 환율 + 해외 이용 수수료).
 *
 * 환율: frankfurter.dev (유럽중앙은행 기준환율, 무료·키 없음). 주말·휴일은 직전 영업일 환율을 돌려준다.
 * 실패하면 null → 금액 칸은 비워 둔 채(기존 동작) 다음 스캔 때 다시 시도.
 */

const cache = new Map<string, number | null>();

/** date(YYYY-MM-DD)의 1 {currency} = ? 원. 실패 시 null */
export async function krwRateOn(currency: string, date: string): Promise<number | null> {
  const key = `${currency}:${date}`;
  if (cache.has(key)) return cache.get(key)!;
  let rate: number | null = null;
  try {
    const res = await fetch(`https://api.frankfurter.dev/v1/${date}?base=${encodeURIComponent(currency)}&symbols=KRW`, {
      signal: AbortSignal.timeout(5000),
      cache: 'no-store',
    });
    if (res.ok) {
      const j = (await res.json()) as { rates?: { KRW?: number } };
      rate = typeof j.rates?.KRW === 'number' && j.rates.KRW > 0 ? j.rates.KRW : null;
    }
  } catch {
    rate = null;
  }
  // 실패는 캐시하지 않는다(같은 실행 안에서도 다음 건이 다시 시도)
  if (rate !== null) cache.set(key, rate);
  return rate;
}

/** 외화 원금 → 추정 원화(원 단위 반올림). 환율 실패 시 null */
export async function estimateKrw(currency: string, original: number, date: string): Promise<number | null> {
  const rate = await krwRateOn(currency, date);
  return rate === null ? null : Math.round(original * rate);
}
