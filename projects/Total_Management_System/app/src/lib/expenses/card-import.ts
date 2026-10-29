/**
 * 카드 이용내역 엑셀 → 경비 「확인 대기」 (2026-10-29)
 *
 * 사장님: 홈택스 「사업용 신용카드 사용내역」(또는 카드사 이용내역) 엑셀을 경비 화면에 올린다.
 * 카드사·홈택스마다 칸 모양이 달라 머리줄(제목 줄)의 글자로 날짜·가맹점·금액 칸을 찾는다.
 * - 취소 줄(음수 금액·「취소」 표시)은 빼고, 같은 가맹점·같은 금액의 승인 줄도 하나 같이 뺀다
 * - 같은 파일을 두 번 올려도 한 번만: gmail_id 자리에 'card:날짜:금액:가맹점:순번' 키 (expense_inbox.gmail_id UNIQUE)
 * - 거래처 기억(expense_vendor_rules)이 메일과 똑같이 적용 — 분류·고정/사용량·증빙·늘 무시
 * 여기서도 expenses 에는 넣지 않는다. 확정은 사장님 [확인] 때만.
 */
import * as XLSX from 'xlsx';
import { defaultsFor } from './mail-parse';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface CardRow { date: string; merchant: string; amount: number; industry: string | null }

const DATE_H = /(승인|거래|이용|사용|매출)\s*(일자|일시|날짜|일)|^일자$|^날짜$/;
const MERCHANT_H = /가맹점\s*명|가맹점|이용\s*처|이용\s*가맹점|상호/;
const AMOUNT_H = [/^합계$|합계\s*금액/, /승인\s*금액/, /이용\s*금액|사용\s*금액|결제\s*금액/, /^금액$|원화\s*금액/];
const INDUSTRY_H = /업종|업태/;
const CANCEL_H = /취소|승인\s*구분|구분|상태/;

function normHead(v: unknown): string { return String(v ?? '').replace(/\s+/g, ' ').trim(); }

function toDate(v: unknown): string | null {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getTime() - v.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  }
  if (typeof v === 'number' && v > 30000 && v < 80000) { // 엑셀 날짜 일련번호
    const d = XLSX.SSF.parse_date_code(v);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const s = String(v ?? '').trim();
  let m = s.match(/(20\d{2})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(20\d{2})(\d{2})(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return null;
}

function toAmount(v: unknown): number | null {
  if (typeof v === 'number') return Math.round(v);
  const s = String(v ?? '').replace(/[,\s원₩]/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Math.round(Number(s));
}

/** 가맹점 이름 정리 — 거래처 기억의 열쇠라 같은 곳은 같은 글자로 */
export function cleanMerchant(s: string): string {
  return s.replace(/\(주\)|㈜|주식회사|\(유\)|유한회사/g, ' ').replace(/\s+/g, ' ').trim();
}

/** 엑셀(xlsx=PK / xls=D0CF)이 아니면 CSV 글자 — 카드사 CSV 는 EUC-KR 이 많아 깨지면 다시 읽는다 */
function readBook(buf: ArrayBuffer | Buffer): XLSX.WorkBook {
  const u8 = buf instanceof ArrayBuffer ? new Uint8Array(buf) : new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  const binary = (u8[0] === 0x50 && u8[1] === 0x4b) || (u8[0] === 0xd0 && u8[1] === 0xcf);
  if (binary) return XLSX.read(u8, { type: 'array', cellDates: true });
  let text = new TextDecoder('utf-8').decode(u8);
  if (text.includes('�')) text = new TextDecoder('euc-kr').decode(u8);
  return XLSX.read(text.replace(/^﻿/, ''), { type: 'string', cellDates: true });
}

export function parseCardSheet(buf: ArrayBuffer | Buffer): { rows: CardRow[]; cancelled: number; error?: string } {
  const wb = readBook(buf);
  for (const name of wb.SheetNames) {
    const aoa = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[name], { header: 1, raw: true, defval: '' });
    // 머리줄 찾기 — 위쪽 30줄 안에 날짜·가맹점·금액 칸 제목이 모두 있는 줄
    for (let h = 0; h < Math.min(30, aoa.length); h++) {
      const head = (aoa[h] || []).map(normHead);
      const di = head.findIndex((c) => DATE_H.test(c));
      const mi = head.findIndex((c) => MERCHANT_H.test(c) && !/번호|사업자/.test(c));
      let ai = -1;
      for (const re of AMOUNT_H) { ai = head.findIndex((c) => re.test(c) && !/공급|부가|봉사|수수료|할부/.test(c)); if (ai >= 0) break; }
      if (di < 0 || mi < 0 || ai < 0) continue;
      const ii = head.findIndex((c) => INDUSTRY_H.test(c));
      const ci = head.findIndex((c, i) => i !== di && i !== mi && i !== ai && CANCEL_H.test(c));

      const ok: CardRow[] = [];
      const cancels: Array<{ merchant: string; amount: number }> = [];
      for (const line of aoa.slice(h + 1)) {
        const date = toDate(line[di]);
        const amount = toAmount(line[ai]);
        const merchant = cleanMerchant(String(line[mi] ?? ''));
        if (!date || amount === null || !merchant) continue; // 합계 줄·빈 줄
        const isCancel = amount < 0 || (ci >= 0 && /취소/.test(String(line[ci])));
        if (isCancel) { cancels.push({ merchant, amount: Math.abs(amount) }); continue; }
        if (amount === 0) continue;
        ok.push({ date, merchant, amount, industry: ii >= 0 ? (String(line[ii] || '').trim() || null) : null });
      }
      // 취소 1건당 같은 가맹점·같은 금액 승인 1건 제거
      for (const c of cancels) {
        const k = ok.findIndex((r) => r.merchant === c.merchant && r.amount === c.amount);
        if (k >= 0) ok.splice(k, 1);
      }
      return { rows: ok, cancelled: cancels.length };
    }
  }
  return { rows: [], cancelled: 0, error: '날짜·가맹점·금액 칸을 찾지 못했습니다 — 카드 이용내역 엑셀이 맞는지 확인해 주세요' };
}

export async function importCardRows(db: Db, rows: CardRow[], fileName: string): Promise<{ added: number; ignored: number; existed: number }> {
  const { data: ruleRows } = await db.from('expense_vendor_rules').select('vendor, category, kind, billing_cycle, proof_type, auto_ignore');
  const rules = new Map<string, { category: string | null; kind: string | null; billing_cycle: string | null; proof_type: string | null; auto_ignore: boolean }>(
    (ruleRows || []).map((r: { vendor: string; category: string | null; kind: string | null; billing_cycle: string | null; proof_type: string | null; auto_ignore: boolean }) => [r.vendor, r]),
  );
  const seq = new Map<string, number>();
  const now = new Date().toISOString();
  const records = rows.map((r) => {
    const base = `card:${r.date}:${r.amount}:${r.merchant}`;
    const n = (seq.get(base) || 0) + 1; // 같은 날 같은 곳 같은 금액 2번 = 순번으로 구분
    seq.set(base, n);
    const rule = rules.get(r.merchant);
    const def = defaultsFor(r.merchant, r.industry || '');
    const kind = rule?.kind || def.kind;
    const ignored = !!rule?.auto_ignore;
    return {
      gmail_id: `${base}:${n}`.slice(0, 300),
      received_at: `${r.date}T12:00:00+09:00`,
      sender: '카드 이용내역 (엑셀)',
      subject: `${r.merchant} · ${fileName}`.slice(0, 300),
      vendor: r.merchant,
      doc_type: 'card',
      amount: r.amount,
      currency: 'KRW',
      amount_original: null,
      expense_date: r.date,
      category: rule?.category || def.category,
      kind,
      billing_cycle: rule?.billing_cycle === 'annual' && kind === 'fixed' ? 'annual' : null,
      proof_type: rule?.proof_type || 'card',
      memo: r.industry,
      snippet: null,
      status: ignored ? 'ignored' : 'pending',
      decided_at: ignored ? now : null,
    };
  });

  let added = 0, ignored = 0;
  for (let i = 0; i < records.length; i += 200) {
    const chunk = records.slice(i, i + 200);
    const { data, error } = await db.from('expense_inbox')
      .upsert(chunk, { onConflict: 'gmail_id', ignoreDuplicates: true }).select('status');
    if (error) throw error;
    for (const d of data || []) { if (d.status === 'ignored') ignored++; else added++; }
  }
  return { added, ignored, existed: records.length - added - ignored };
}
