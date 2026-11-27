/**
 * 한국 시간(KST) 날짜 — 날짜를 다룰 때는 이 파일의 함수만 쓴다 (2026-11-15)
 *
 * 왜: `new Date().toISOString().slice(0, 10)` 은 UTC 날짜다. 한국 자정~오전 9시에는 "어제"가 나온다.
 *   서버(Vercel)는 항상 UTC 라서 오전 9시 전에 만든 판매·발주의 날짜가 하루 앞으로 찍혔고,
 *   화면에서는 「마지막 거래 -1일 전」 같은 값이 나왔다. 71곳에 흩어져 있던 것을 여기로 모은다.
 *
 * 모든 함수는 실행 환경의 시간대(서버 UTC / 브라우저 KST)와 무관하게 같은 값을 낸다.
 */
const KST_OFFSET_MS = 9 * 60 * 60 * 1000;
const DAY_MS = 24 * 60 * 60 * 1000;

/** 어떤 시각(Date · ISO 문자열 · ms)이 한국에서 며칠인지 → 'YYYY-MM-DD' */
export function kstDate(input: Date | string | number): string {
  const t = input instanceof Date ? input.getTime() : typeof input === 'number' ? input : new Date(input).getTime();
  if (!Number.isFinite(t)) return '';
  return new Date(t + KST_OFFSET_MS).toISOString().slice(0, 10);
}

/**
 * 고객 화면용 한국 날짜 'YYYY. M. D.' (toLocaleDateString('ko-KR') 과 같은 모양) — 2026-11-27
 * 서버(UTC)에서 toLocaleDateString 을 그냥 부르면 오전 9시 전 시각이 하루 전 날짜로 나온다. 값이 없으면 ''.
 */
export function kstDateDots(input: Date | string | number | null | undefined): string {
  if (input == null || input === '') return '';
  const s = kstDate(input);
  if (!s) return '';
  const [y, m, d] = s.split('-').map(Number);
  return `${y}. ${m}. ${d}.`;
}

/** 오늘(한국) 'YYYY-MM-DD' */
export function todayKST(): string {
  return kstDate(Date.now());
}

/** 'YYYY-MM-DD' 에 n일 더하기(음수 가능) → 'YYYY-MM-DD'. 날짜만 계산하므로 시간대와 무관 */
export function addDays(dateStr: string, n: number): string {
  const [y, m, d] = dateStr.slice(0, 10).split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + n)).toISOString().slice(0, 10);
}

/** 'YYYY-MM-DD' 에 n개월 더하기(음수 가능). 말일을 넘으면 그 달 말일로 맞춘다 */
export function addMonths(dateStr: string, n: number): string {
  const [y, m, d] = dateStr.slice(0, 10).split('-').map(Number);
  const last = new Date(Date.UTC(y, m - 1 + n + 1, 0)).getUTCDate();
  return new Date(Date.UTC(y, m - 1 + n, Math.min(d, last))).toISOString().slice(0, 10);
}

/** 이번 달 1일(한국) 'YYYY-MM-01' */
export function monthStartKST(): string {
  return todayKST().slice(0, 8) + '01';
}

/** 그 날짜(또는 시각)로부터 오늘(한국)까지 며칠 지났나. 오늘이면 0, 미래면 음수 */
export function daysSinceKST(input: string | Date | number): number {
  const s = typeof input === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(input) ? input : kstDate(input);
  if (!s) return NaN;
  const [y, m, d] = s.split('-').map(Number);
  const [ty, tm, td] = todayKST().split('-').map(Number);
  return Math.round((Date.UTC(ty, tm - 1, td) - Date.UTC(y, m - 1, d)) / DAY_MS);
}
