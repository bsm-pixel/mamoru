/**
 * 우리WON뱅킹 입출금 알림 문구 해석 (161)
 *
 * 실측 형식 (2026-10-13 사장님 휴대폰 알림):
 *   제목: 우리WON뱅킹 입출금알림
 *   본문: [입금] 민세영      220,000원 1002-439-462***계좌 10/13 14:52:18
 *
 * [입금] 만 해석한다. [출금]·광고·기타 알림은 null.
 */
export interface ParsedDeposit {
  depositor: string;     // 입금자명 (은행 표기 그대로, 앞뒤 공백 제거)
  amount: number;        // 원
  depositedAt: string;   // ISO (KST 기준 시각)
  dedupeKey: string;     // 같은 알림이 두 번 와도 한 번만 처리
}

const DEPOSIT_RE = /\[입금\]\s*(.+?)\s+([\d,]+)\s*원/;
const TIME_RE = /(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/;

export function parseWooriDeposit(text: string, now: Date = new Date()): ParsedDeposit | null {
  const flat = text.replace(/\s+/g, ' ').trim();
  const m = flat.match(DEPOSIT_RE);
  if (!m) return null;

  const depositor = m[1].trim();
  const amount = Number(m[2].replace(/,/g, ''));
  if (!depositor || !Number.isFinite(amount) || amount <= 0) return null;

  // 알림엔 연도가 없다 → KST 현재 연도. 1월에 12월 알림이 오면 작년
  const kstNow = new Date(now.getTime() + 9 * 3600_000);
  let year = kstNow.getUTCFullYear();
  let depositedAt = now.toISOString();
  let stamp = '';
  const t = flat.match(TIME_RE);
  if (t) {
    const [, mo, d, hh, mi, ss = '00'] = t;
    if (Number(mo) > kstNow.getUTCMonth() + 1) year -= 1;
    const p = (v: string) => v.padStart(2, '0');
    const iso = `${year}-${p(mo)}-${p(d)}T${p(hh)}:${mi}:${ss}+09:00`;
    const parsed = new Date(iso);
    if (!Number.isNaN(parsed.getTime())) {
      depositedAt = parsed.toISOString();
      stamp = iso;
    }
  }

  return {
    depositor,
    amount,
    depositedAt,
    // 시각이 없으면 문구 전체로 구분 (같은 문구 재전송만 막힘)
    dedupeKey: `woori|${depositor}|${amount}|${stamp || flat}`,
  };
}

/** 이름 비교용 정규화 — 공백만 제거 (은행 표기는 공백 포함일 수 있음) */
export function normalizeName(name: string | null | undefined): string {
  return (name || '').replace(/\s+/g, '');
}
