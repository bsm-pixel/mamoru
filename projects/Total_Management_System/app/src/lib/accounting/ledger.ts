/**
 * 종소세 간편장부 (2026-10-29) — GET /api/accounting/ledger?year= 가 엑셀로 내려준다.
 *
 * 사장님: 간이과세자·직접 신고. 1년치 수입·비용을 거래 한 줄씩(간편장부 형식) + 월별 요약 + 증빙 챙길 것.
 *
 * 매출 규칙 = 회계 화면(api/reports/summary)과 같다. 단 세무용이라 두 가지만 다르다:
 *  ① 아임웹 주문 = 실결제액(paid_amount: 배송비 포함·적립금/쿠폰 차감). 회계 화면은 상품 금액(배송비 제외).
 *     → 차이는 「안내」 시트에 금액으로 보여 준다 (reconcile.onlineDiff).
 *  ② 날짜 경계 = 한국시간(+09:00). 회계 화면은 시각 칸(ordered_at·paid_at·returned_at)을 UTC 로 잘라
 *     1/1 00~09시 거래가 앞 해로 갈 수 있다.
 * 금액 공식 SSOT: lib/sales/amounts.ts — offline_sales=GROSS(total−discount), deliveries=NET(그대로).
 * 비용 = 매입(purchase_orders, 원화 저장) + 경비(expenses). 연간 결제는 결제일 전액(현금주의).
 */
import { normalizeImwebPayMethod } from '@/lib/sales/imweb-pay-method';
import { proofLabel, proofNeedsKeeping } from '@/lib/expenses/proof';
import { planLabel } from '@/lib/expenses/plan';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

export interface LedgerRow {
  date: string;          // YYYY-MM-DD (한국시간)
  kind: '매출' | '매출(반품)' | '매입' | '경비';
  content: string;       // 거래내용
  party: string;         // 거래처
  income: number;        // 수입
  cost: number;          // 비용
  proof: string;         // 결제·증빙
  note: string;          // 비고
  keep?: string;         // 「증빙 챙길 것」 사유 (있으면 그 시트에 올라감)
}

export interface Ledger {
  year: number;
  rows: LedgerRow[];
  months: Array<{ month: string; sales: number; purchases: number; expenses: number }>;
  totals: { income: number; purchases: number; expenses: number };
  reconcile: { onlinePaid: number; onlineItems: number; onlineDiff: number; reportStyleIncome: number };
}

const METHOD: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합', point: '적립금', phone: '휴대폰결제', unknown: '',
};
function methodLabel(m: string | null | undefined): string {
  const k = (m || '').toLowerCase();
  if (k in METHOD) return METHOD[k];
  if (/pay/.test(k)) return `간편결제(${k.replace(/direct$/, '')})`; // kakaopaydirect·naverpay 등 = 카드매출과 같은 PG 결제
  return m || '';
}
function mixedLabel(d: Record<string, number> | null): string {
  if (!d) return '복합';
  return (['card', 'cash', 'transfer'] as const).filter((k) => d[k]).map((k) => `${METHOD[k]} ${d[k].toLocaleString()}`).join(' / ') || '복합';
}
/** timestamptz → 한국 날짜 */
function kstDate(iso: string): string {
  return new Date(new Date(iso).getTime() + 9 * 3600000).toISOString().slice(0, 10);
}
function itemsText(names: string[]): string {
  const n = names.filter((x) => x && x !== '배송비');
  if (!n.length) return '';
  return n.length === 1 ? n[0] : `${n[0]} 외 ${n.length - 1}`;
}

async function itemsBy(db: Db, table: string, key: string, ids: string[]): Promise<Record<string, Array<{ product_name: string; total_price: number; category?: string | null }>>> {
  const out: Record<string, Array<{ product_name: string; total_price: number; category?: string | null }>> = {};
  for (let i = 0; i < ids.length; i += 200) {
    const cols = table === 'order_items' ? `${key}, product_name, total_price, unit_price, quantity` : `${key}, product_name, total_price, category`;
    const { data, error } = await db.from(table).select(cols).in(key, ids.slice(i, i + 200));
    if (error) throw error;
    for (const it of data || []) {
      const total = it.total_price || (it.unit_price || 0) * (it.quantity || 0);
      (out[it[key]] ||= []).push({ product_name: it.product_name, total_price: total, category: it.category });
    }
  }
  return out;
}

export async function buildLedger(db: Db, year: number): Promise<Ledger> {
  const from = `${year}-01-01`;
  const to = `${year}-12-31`;
  const tsFrom = `${year}-01-01T00:00:00+09:00`;
  const tsTo = `${year + 1}-01-01T00:00:00+09:00`;
  const rows: LedgerRow[] = [];

  // ── 1) 판매 (판매일, total−discount, RS 포함) ──
  const { data: sales, error: e1 } = await db.from('offline_sales')
    .select('id, sale_number, sale_date, customer_name, total_amount, discount_amount, payment_method, payment_detail, payment_status')
    .gte('sale_date', from).lte('sale_date', to).is('cancelled_at', null);
  if (e1) throw e1;
  // ── 1-2) 반품 (반품한 날 −) ──
  const { data: rets, error: e2 } = await db.from('offline_sales')
    .select('id, sale_number, sale_date, returned_at, customer_name, total_amount, discount_amount')
    .gte('returned_at', tsFrom).lt('returned_at', tsTo).is('cancelled_at', null);
  if (e2) throw e2;
  const saleItems = await itemsBy(db, 'offline_sale_items', 'sale_id', [...new Set([...(sales || []), ...(rets || [])].map((s: { id: string }) => s.id))]);
  for (const s of sales || []) {
    const its = saleItems[s.id] || [];
    const hasRs = its.some((it) => it.category === 'RS' && it.total_price > 0);
    rows.push({
      date: s.sale_date, kind: '매출',
      content: ['판매', itemsText(its.map((i) => i.product_name))].filter(Boolean).join(' · '),
      party: s.customer_name || '', income: (s.total_amount || 0) - (s.discount_amount || 0), cost: 0,
      proof: s.payment_method === 'mixed' ? mixedLabel(s.payment_detail) : methodLabel(s.payment_method),
      note: [s.sale_number, hasRs ? '복원수리 포함' : '', s.payment_status !== 'paid' ? '미수 있음' : ''].filter(Boolean).join(' · '),
    });
  }
  for (const s of rets || []) {
    rows.push({
      date: kstDate(s.returned_at), kind: '매출(반품)',
      content: ['반품', itemsText((saleItems[s.id] || []).map((i) => i.product_name))].filter(Boolean).join(' · '),
      party: s.customer_name || '', income: -((s.total_amount || 0) - (s.discount_amount || 0)), cost: 0,
      proof: '', note: `${s.sale_number} (판매일 ${s.sale_date})`,
    });
  }

  // ── 2) 납품 (납품일, total 그대로 = 이미 NET) ──
  const { data: dls, error: e3 } = await db.from('deliveries')
    .select('id, dl_number, delivery_date, customer_name, total_amount, payment_method, payment_status, receipt_type')
    .gte('delivery_date', from).lte('delivery_date', to)
    .in('status', ['confirmed', 'shipped', 'settled']).is('cancelled_at', null);
  if (e3) throw e3;
  const dlItems = await itemsBy(db, 'delivery_items', 'delivery_id', (dls || []).map((d: { id: string }) => d.id));
  for (const d of dls || []) {
    rows.push({
      date: d.delivery_date, kind: '매출',
      content: ['납품', itemsText((dlItems[d.id] || []).map((i) => i.product_name))].filter(Boolean).join(' · '),
      party: d.customer_name || '', income: d.total_amount || 0, cost: 0,
      proof: methodLabel(d.payment_method),
      note: [d.dl_number, d.payment_status !== 'paid' ? '미수 있음' : ''].filter(Boolean).join(' · '),
    });
  }

  // ── 3) 아임웹 주문 (주문일 한국시간, 실결제액) ──
  const { data: ords, error: e4 } = await db.from('orders')
    .select('id, imweb_order_no, ordered_at, orderer_name, paid_amount, delivery_fee, payment_method')
    .gte('ordered_at', tsFrom).lt('ordered_at', tsTo)
    .not('status', 'in', '("cancelled","refunded","pay_wait")');
  if (e4) throw e4;
  const ordItems = await itemsBy(db, 'order_items', 'order_id', (ords || []).map((o: { id: string }) => o.id));
  let onlinePaid = 0, onlineItems = 0;
  for (const o of ords || []) {
    const its = ordItems[o.id] || [];
    const goods = its.filter((i) => i.product_name !== '배송비' && i.total_price > 0).reduce((s, i) => s + i.total_price, 0);
    onlinePaid += o.paid_amount || 0;
    onlineItems += goods;
    if (!(o.paid_amount > 0)) continue;
    const diff = (o.paid_amount || 0) - goods;
    rows.push({
      date: kstDate(o.ordered_at), kind: '매출',
      content: ['아임웹 주문', itemsText(its.map((i) => i.product_name))].filter(Boolean).join(' · '),
      party: o.orderer_name || '', income: o.paid_amount || 0, cost: 0,
      proof: methodLabel(normalizeImwebPayMethod(o.payment_method)),
      note: [o.imweb_order_no, o.delivery_fee > 0 ? `배송비 ${o.delivery_fee.toLocaleString()} 포함` : '', diff < 0 ? `적립금·쿠폰 ${(-diff).toLocaleString()} 차감` : ''].filter(Boolean).join(' · '),
    });
  }

  // ── 4) 복원수리 접수 (입금일 한국시간, 판매와 연결된 건 제외 — 판매 행에 이미 포함) ──
  const { data: reps, error: e5 } = await db.from('repairs')
    .select('as_id, name, total_amount, paid_at, payment_method')
    .not('paid_at', 'is', null).gte('paid_at', tsFrom).lt('paid_at', tsTo)
    .not('status', 'eq', 'cancelled').is('sale_id', null);
  if (e5) throw e5;
  for (const r of reps || []) {
    if (!(r.total_amount > 0)) continue;
    rows.push({
      date: kstDate(r.paid_at), kind: '매출', content: '복원수리', party: r.name || '',
      income: r.total_amount, cost: 0, proof: methodLabel(r.payment_method), note: r.as_id || '',
    });
  }

  // ── 5) 매입 (발주일, 원화) ──
  const { data: pos, error: e6 } = await db.from('purchase_orders')
    .select('po_number, order_date, supplier_name, total_amount, vat_amount, currency, exchange_rate, status')
    .gte('order_date', from).lte('order_date', to)
    .in('status', ['received', 'balance_paid', 'deposit_paid', 'ordered']);
  if (e6) throw e6;
  for (const p of pos || []) {
    const fx = p.currency && p.currency !== 'KRW' ? `${p.currency} 환율 ${p.exchange_rate}` : '';
    const taxInv = (p.vat_amount || 0) > 0;
    rows.push({
      date: p.order_date, kind: '매입', content: '상품 매입', party: p.supplier_name || '',
      income: 0, cost: p.total_amount || 0,
      proof: taxInv ? '세금계산서' : (fx ? '해외 송금' : ''),
      note: [p.po_number, fx, p.status === 'ordered' ? '아직 미지급' : ''].filter(Boolean).join(' · '),
      keep: taxInv ? undefined : (fx ? '해외 매입 — 인보이스·송금 내역 보관' : '세금계산서 없음 — 거래명세서·이체 내역 보관'),
    });
  }

  // ── 6) 경비 ──
  const { data: exps, error: e7 } = await db.from('expenses')
    .select('expense_date, category, amount, memo, vendor, kind, billing_cycle, proof_type, source_ref')
    .gte('expense_date', from).lte('expense_date', to);
  if (e7) throw e7;
  for (const e of exps || []) {
    const plan = planLabel(e.kind, e.billing_cycle);
    rows.push({
      date: e.expense_date, kind: '경비',
      content: [e.category, e.memo && e.memo !== e.vendor ? e.memo : ''].filter(Boolean).join(' · '),
      party: e.vendor || '', income: 0, cost: e.amount || 0,
      proof: e.proof_type ? proofLabel(e.proof_type) : '미정',
      note: [plan ? `${plan}${plan === '연간' ? ' (결제일 전액)' : ''}` : '', e.source_ref?.startsWith('mail:') ? '메일 영수증' : e.source_ref?.startsWith('card:') ? '카드 내역' : ''].filter(Boolean).join(' · '),
      keep: proofNeedsKeeping(e.proof_type)
        ? (!e.proof_type ? '증빙 미정 — 경비 화면에서 골라 주세요' : e.proof_type === 'receipt' ? (e.source_ref?.startsWith('mail:') ? '영수증만 — 메일 원본 보관' : '영수증만 — 원본 보관') : '증빙 없음 — 이체 내역 등 보관')
        : undefined,
    });
  }

  rows.sort((a, b) => a.date.localeCompare(b.date) || a.kind.localeCompare(b.kind));

  const months = Array.from({ length: 12 }, (_, i) => ({ month: `${year}-${String(i + 1).padStart(2, '0')}`, sales: 0, purchases: 0, expenses: 0 }));
  for (const r of rows) {
    const m = months[parseInt(r.date.slice(5, 7), 10) - 1];
    if (!m) continue;
    m.sales += r.income;
    if (r.kind === '매입') m.purchases += r.cost; else m.expenses += r.cost;
  }
  const totals = {
    income: months.reduce((s, m) => s + m.sales, 0),
    purchases: months.reduce((s, m) => s + m.purchases, 0),
    expenses: months.reduce((s, m) => s + m.expenses, 0),
  };
  const onlineDiff = onlinePaid - onlineItems;
  return { year, rows, months, totals, reconcile: { onlinePaid, onlineItems, onlineDiff, reportStyleIncome: totals.income - onlineDiff } };
}
