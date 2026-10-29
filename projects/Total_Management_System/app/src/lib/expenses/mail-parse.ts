/**
 * 영수증·세금계산서 메일 1통 → 경비 후보 (2026-10-29)
 *
 * 순수 함수(DB·네트워크 없음) — 규칙으로만 읽는다. Claude API 같은 유료 판독은 쓰지 않는다(CLAUDE.md 5.5).
 * 못 읽는 메일은 null(목록에 안 올림). 금액이 PDF에만 있는 메일(구글 워크스페이스·솔라피 거래명세서·일부 세금계산서)은
 * 금액 없이 올려 사장님이 확인 때 입력한다.
 *
 * 2026-10-29 실제 메일로 확인한 형식:
 *   KG파이낸셜  "결제금액 | 30,500원" · "결제일시 | 2026-10-11 10:46:15" · "구매상점 | (주)프로스퍼글로벌"
 *   NHN KCP     "Adobe에서 2026년 10월 07일 … 결제하신" · "결제금액 | 39,600 원"
 *   Google Play "Google Play의 Arlosoft Ltd에서 구매" · "합계: ₩10,000"
 *   Anthropic   "Receipt from Anthropic, PBC $100.00 Paid on October 26, 2026" (외화 → 원화는 확인 때)
 *   팝빌        "[(주)비바리퍼블리카]에서 전자세금계산서" · "공급가액 | -1,280 원" · "세액 | -128 원"
 *   유니포스트  "주식회사 카카오페이(으)로 부터 세금계산서" (금액 없음)
 */

export type DocType = 'card' | 'receipt' | 'invoice' | 'tax_invoice';
export type ExpenseKind = 'fixed' | 'variable';

export interface ParsedExpense {
  vendor: string;
  docType: DocType;
  amount: number | null;          // 원화
  currency: 'KRW' | 'USD';
  amountOriginal: number | null;  // 외화 원금
  expenseDate: string;            // YYYY-MM-DD (KST)
  memo: string;
  // 처음 보는 거래처일 때의 기본값 — 사장님이 [확인] 때 고른 값이 거래처 규칙으로 저장되면 그쪽이 이긴다
  category: string;
  kind: ExpenseKind;
}

export interface MailInput {
  from: string;
  subject: string;
  text: string;
  receivedAt: string; // ISO
}

/** 영수증이 아닌 메일 — 보낸 곳·제목만으로 거른다 (공지·광고·검수 알림·고객 주문 알림 등) */
const SKIP_SUBJECT = /작업\s*공지|공지|업데이트|개인정보|입금계좌|검수|약관|소식|뉴스레터|이벤트|프로모션|\[MAMORU\]|새 주문|주문이 접수|주문 알림/;
const SKIP_SENDER = /imweb\.me|danal\.co\.kr|payaction\.app|wallet-email\.samsung\.com|bsm@mamoru\.kr/i;

/** 거래처별 기본 분류 — 처음 한 번만 쓰이고, 이후엔 사장님이 고른 값(expense_vendor_rules)이 우선 */
const VENDOR_DEFAULTS: Array<{ re: RegExp; category: string; kind: ExpenseKind }> = [
  { re: /anthropic|claude/i, category: '소프트웨어', kind: 'fixed' },
  { re: /adobe/i, category: '소프트웨어', kind: 'fixed' },
  { re: /google workspace/i, category: '소프트웨어', kind: 'fixed' },
  { re: /notion|figma|vercel|supabase|openai|make\.com|jitter/i, category: '소프트웨어', kind: 'fixed' },
  { re: /솔라피|solapi/i, category: '소프트웨어', kind: 'variable' },
  { re: /google play|arlosoft/i, category: '소프트웨어', kind: 'variable' },
  { re: /롯데|lotte|택배|우체국|cj대한통운|postteam|프로스퍼글로벌|배송비/i, category: '택배비', kind: 'variable' },
  { re: /쿠팡|coupang|1688|알리|aliexpress|타오바오|포장|박스/i, category: '포장재', kind: 'variable' },
  { re: /임대|관리비|건물/i, category: '임대료', kind: 'fixed' },
];

export function defaultsFor(vendor: string, memo = ''): { category: string; kind: ExpenseKind } {
  const hay = `${vendor} ${memo}`;
  for (const d of VENDOR_DEFAULTS) if (d.re.test(hay)) return { category: d.category, kind: d.kind };
  return { category: '기타', kind: 'variable' };
}

const kstDate = (iso: string) => new Date(new Date(iso).getTime() + 9 * 3600 * 1000).toISOString().slice(0, 10);
const num = (s: string) => parseInt(s.replace(/[,\s]/g, ''), 10);
const pad = (n: string | number) => String(n).padStart(2, '0');
const clean = (s: string) => s.replace(/\s+/g, ' ').replace(/^[\s|:]+|[\s|:]+$/g, '').trim();

/** "결제금액 | 30,500원" 처럼 라벨 뒤 첫 금액(원). 음수(취소·수정 세금계산서) 허용 */
function amountAfter(text: string, labels: string): number | null {
  const m = new RegExp(`(?:${labels})[\\s|:：]*(-?[\\d,]{1,12})\\s*원`).exec(text);
  return m ? num(m[1]) : null;
}

/** "2026-10-11 10:46" · "2026년 10월 07일" · "2026. 10. 20." · "20260930" → YYYY-MM-DD */
function dateAfter(text: string, labels: string): string | null {
  const m = new RegExp(`(?:${labels})[\\s|:：]*(\\d{4})\\s*[-.년]?\\s*(\\d{1,2})\\s*[-.월]?\\s*(\\d{1,2})`).exec(text);
  return m ? `${m[1]}-${pad(m[2])}-${pad(m[3])}` : null;
}

function senderAddr(from: string): string {
  const m = /<([^>]+)>/.exec(from);
  return (m ? m[1] : from).trim().toLowerCase();
}

export function parseExpenseMail(mail: MailInput): ParsedExpense | null {
  const { subject, text } = mail;
  const addr = senderAddr(mail.from);
  if (SKIP_SENDER.test(addr) || SKIP_SUBJECT.test(subject)) return null;

  const received = kstDate(mail.receivedAt);
  const t = text.replace(/ /g, ' ');
  let vendor = '';
  let docType: DocType = 'receipt';
  let amount: number | null = null;
  let currency: 'KRW' | 'USD' = 'KRW';
  let amountOriginal: number | null = null;
  let date: string | null = null;
  let memo = '';

  if (/kgfinancial|kgmobilians/.test(addr)) {
    // KG파이낸셜(구 KG모빌리언스) 카드 결제
    docType = 'card';
    vendor = clean(/구매상점[\s|]*([^|\n]+)/.exec(t)?.[1] || '') || clean(/구매사이트[\s|]*([^|\n]+)/.exec(t)?.[1] || '');
    memo = clean(/상품명[\s|]*([^|\n]+)/.exec(t)?.[1] || '');
    amount = amountAfter(t, '결제금액');
    date = dateAfter(t, '결제일시');
  } else if (/kcp\.co\.kr/.test(addr)) {
    // NHN KCP — "Adobe에서 2026년 10월 07일 … 결제하신"
    docType = 'card';
    vendor = clean(/구매상점명[\s|]*([^|\n]+)/.exec(t)?.[1] || '') || clean(/^NHN KCP\s*-\s*(.+?)의 결제/.exec(subject)?.[1] || '');
    memo = clean(/주문상품명[\s|]*([^|\n]+)/.exec(t)?.[1] || '');
    amount = amountAfter(t, '결제금액');
    date = dateAfter(t, '결제일시');
  } else if (/kiwoompay/.test(addr)) {
    docType = 'card';
    vendor = clean(/님의\s*(.+?)\s*결제\s*내역/.exec(subject)?.[1] || '');
    amount = amountAfter(t, '결제금액|승인금액|결제 금액');
    date = dateAfter(t, '결제일시|승인일시|거래일시');
  } else if (/googleplay-noreply@google\.com/.test(addr)) {
    docType = 'receipt';
    vendor = `Google Play · ${clean(/Google Play의\s*(.+?)에서 구매/.exec(t)?.[1] || '')}`.replace(/ · $/, '');
    memo = clean(/상품 가격\s*\n?\s*([^₩\n]+)₩/.exec(t)?.[1] || '');
    const m = /합계[:\s]*₩\s*([\d,]+)/.exec(t);
    amount = m ? num(m[1]) : null;
    date = /주문 날짜:\s*(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})/.exec(t)?.slice(1, 4).map(pad).join('-') || null;
  } else if (/payments-noreply@google\.com/.test(addr) && /workspace/i.test(subject)) {
    docType = 'invoice';
    vendor = 'Google Workspace';
    memo = clean(/인보이스 번호\s*([\d]+)/.exec(t)?.[0] || '월별 인보이스 (금액은 첨부 PDF)');
  } else if (/solapi\.com/.test(addr) && /거래명세서/.test(subject)) {
    docType = 'invoice';
    vendor = '솔라피';
    memo = '거래명세서 (금액은 첨부 PDF)';
  } else if (/세금계산서/.test(subject)) {
    // 팝빌·유니포스트·홈택스 등 — 공급자 이름은 제목에서
    docType = 'tax_invoice';
    vendor = clean(/^\[(.+?)\]에서/.exec(subject)?.[1] || /님\s*(.+?)\s*\(으\)로\s*부터/.exec(subject)?.[1] || /(.+?)\s*\(으\)로\s*부터/.exec(subject)?.[1] || '');
    if (!vendor) vendor = clean(/공급자[\s\S]{0,40}?상호[\s|]*([^|\n]+)/.exec(t)?.[1] || '');
    const supply = amountAfter(t, '공급가액');
    const tax = amountAfter(t, '세액');
    const total = amountAfter(t, '합계금액|합계 금액|총액');
    amount = total ?? (supply !== null ? supply + (tax || 0) : null);
    date = dateAfter(t, '작성일자');
    memo = '전자세금계산서';
  } else {
    // 그 밖의 영수증 — Stripe/Paddle 계열 달러 영수증, 원화 결제 안내
    const usd = /\$\s?([\d,]+\.\d{2})\s*(?:Paid|paid|USD)?/.exec(t);
    const looksReceipt = /receipt|invoice|영수증|결제\s*(완료|내역|확인)|결제가 완료|청구/i.test(`${subject} ${t.slice(0, 400)}`);
    if (!looksReceipt) return null;
    vendor = clean(/Receipt from\s+(.+?)\s+\$/i.exec(t)?.[1] || '')
      || clean(/^\[([^\]]+)\]/.exec(subject)?.[1] || '')
      || clean((/^"?([^"<]+)"?\s*</.exec(mail.from)?.[1] || addr.split('@')[1] || '').replace(/\s*(Inc|PBC|LLC)\.?$/i, ''));
    const krw = amountAfter(t, '결제\\s*금액|결제금액|총\\s*결제\\s*금액|합계|총액|승인금액|Total');
    if (krw !== null) amount = krw;
    else if (usd) { currency = 'USD'; amountOriginal = parseFloat(usd[1].replace(/,/g, '')); }
    const en = /Paid on\s+([A-Z][a-z]+)\s+(\d{1,2}),\s+(\d{4})/.exec(t);
    if (en) {
      const mon = new Date(`${en[1]} 1, 2000`).getMonth() + 1;
      if (mon > 0) date = `${en[3]}-${pad(mon)}-${pad(en[2])}`;
    }
    date = date || dateAfter(t, '결제일시|결제일|승인일시');
    memo = clean(/(Max plan[^$\n]*|Pro plan[^$\n]*|Team plan[^$\n]*)/i.exec(t)?.[1] || '') || clean(subject).slice(0, 80);
    docType = 'receipt';
    // 금액도 외화도 못 찾은 "결제" 메일은 영수증으로 보기 어렵다
    if (amount === null && amountOriginal === null) return null;
  }

  if (!vendor) vendor = addr.split('@')[1] || addr;
  const d = defaultsFor(vendor, memo);
  return {
    vendor,
    docType,
    amount,
    currency,
    amountOriginal,
    expenseDate: date && /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : received,
    memo,
    category: d.category,
    kind: d.kind,
  };
}

/** Gmail 검색어 — 영수증·세금계산서 후보만 (최종 판정은 parseExpenseMail) */
export function expenseMailQuery(days: number): string {
  return `newer_than:${days}d (영수증 OR receipt OR invoice OR 인보이스 OR 세금계산서 OR 거래명세서 OR "결제 내역" OR 결제내역 OR "결제 확인" OR "결제가 완료" OR 결제완료 OR 청구서) -from:bsm@mamoru.kr`;
}
