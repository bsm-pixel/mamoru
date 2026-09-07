'use client';

import { useState, useEffect, useMemo } from 'react';
import { Printer } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';
import { repairSlip, saleSlip, wrapTray, buildListDoc, esc, type PrepInspection, type PrepListRow } from '@/lib/prep/tray';

/**
 * 통합 준비표 — 복원수리(출고대기)·주문(배송대기)·판매(미출고)를 탭으로 모아 체크 → 한 번에 인쇄.
 * 진입: 세 목록의 '통합 준비표' 버튼 → initialTab 포커스 + 현재 선택분 preselect. (2026-09-07)
 * 체크는 인쇄 대상 선택일 뿐 데이터 무변경(읽기전용).
 */

type Domain = 'repair' | 'order' | 'sale';
interface Row { id: string; label: string; name: string; sub?: string; hasInvoice: boolean }
interface Props {
  initialTab?: Domain;
  preselect?: Partial<Record<Domain, string[]>>;
  onClose: () => void;
}

const TABS: { key: Domain; label: string }[] = [
  { key: 'repair', label: '복원수리' },
  { key: 'order', label: '주문' },
  { key: 'sale', label: '판매' },
];

export function UnifiedPrepModal({ initialTab = 'repair', preselect, onClose }: Props) {
  const [tab, setTab] = useState<Domain>(initialTab);
  const [rows, setRows] = useState<Record<Domain, Row[]>>({ repair: [], order: [], sale: [] });
  const [checked, setChecked] = useState<Record<Domain, Set<string>>>({ repair: new Set(), order: new Set(), sale: new Set() });
  const [loading, setLoading] = useState(true);
  const [printing, setPrinting] = useState(false);
  const [mode, setMode] = useState<'tray' | 'list'>('tray');

  // 준비 대기 목록 로딩(요약만) + preselect 초기화
  useEffect(() => {
    (async () => {
      setLoading(true);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = createClient() as any;
      const [repRes, ordRes, saleRes] = await Promise.all([
        db.from('repairs').select('id, as_id, name, qty_mamoru, qty_other, invoice_number')
          .eq('status', 'ready_to_ship').order('received_at', { ascending: true }).limit(100),
        db.from('orders').select('id, imweb_order_no, recipient_name, orderer_name, invoice_number')
          .eq('status', 'ready_to_ship').order('ordered_at', { ascending: true }).limit(100),
        db.from('offline_sales').select('id, sale_number, customer_name, invoice_number')
          .eq('payment_status', 'paid').is('shipped_at', null).is('delivered_at', null)
          .is('cancelled_at', null).is('returned_at', null).order('sale_date', { ascending: true }).limit(100),
      ]);
      const repair: Row[] = (repRes.data || []).map((r: Record<string, unknown>) => ({
        id: r.id as string, label: (r.as_id as string) || '', name: (r.name as string) || '이름없음',
        sub: `마모루 ${r.qty_mamoru || 0}자루${(r.qty_other as number) > 0 ? ` · 타사 ${r.qty_other}자루` : ''}`,
        hasInvoice: !!r.invoice_number,
      }));
      const order: Row[] = (ordRes.data || []).map((o: Record<string, unknown>) => ({
        id: o.id as string, label: (o.imweb_order_no as string) || '', name: (o.recipient_name as string) || (o.orderer_name as string) || '',
        hasInvoice: !!o.invoice_number,
      }));
      const sale: Row[] = (saleRes.data || []).map((s: Record<string, unknown>) => ({
        id: s.id as string, label: (s.sale_number as string) || '', name: (s.customer_name as string) || '',
        hasInvoice: !!s.invoice_number,
      }));
      setRows({ repair, order, sale });
      // preselect: 넘어온 id 중 목록에 있는 것만 체크
      const init: Record<Domain, Set<string>> = { repair: new Set(), order: new Set(), sale: new Set() };
      (['repair', 'order', 'sale'] as Domain[]).forEach((d) => {
        const ids = new Set((d === 'repair' ? repair : d === 'order' ? order : sale).map((x) => x.id));
        (preselect?.[d] || []).forEach((id) => { if (ids.has(id)) init[d].add(id); });
      });
      setChecked(init);
      setLoading(false);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const toggle = (d: Domain, id: string) => setChecked((prev) => {
    const next = { ...prev, [d]: new Set(prev[d]) };
    if (next[d].has(id)) next[d].delete(id); else next[d].add(id);
    return next;
  });
  const toggleAll = (d: Domain) => setChecked((prev) => {
    const all = rows[d].map((r) => r.id);
    const cur = prev[d];
    const next = new Set<string>(cur.size === all.length ? [] : all);
    return { ...prev, [d]: next };
  });

  const totalChecked = checked.repair.size + checked.order.size + checked.sale.size;

  const handlePrint = async () => {
    if (totalChecked === 0) return;
    setPrinting(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = createClient() as any;
      const slips: string[] = [];
      const listRows: PrepListRow[] = [];

      // 1) 복원수리 슬립
      const repIds = [...checked.repair];
      if (repIds.length) {
        const [{ data: reps }, { data: insp }] = await Promise.all([
          db.from('repairs').select('*').in('id', repIds),
          db.from('repair_inspections').select('repair_id, scissor_number, scissor_type, comment').in('repair_id', repIds).order('scissor_number'),
        ]);
        const inspBy: Record<string, PrepInspection[]> = {};
        (insp || []).forEach((i: PrepInspection & { repair_id: string }) => { (inspBy[i.repair_id] ||= []).push(i); });
        const custIds = [...new Set((reps || []).map((r: Record<string, unknown>) => r.customer_id).filter(Boolean))] as string[];
        const actBy: Record<string, string> = {};
        if (custIds.length) {
          const { data: custs } = await db.from('customers').select('id, activity_name').in('id', custIds);
          (custs || []).forEach((c: { id: string; activity_name?: string | null }) => { if (c.activity_name) actBy[c.id] = c.activity_name; });
        }
        const ordered = repIds.map((id) => (reps || []).find((r: Record<string, unknown>) => r.id === id)).filter(Boolean);
        ordered.forEach((r: Record<string, unknown>) => {
          const insp2 = inspBy[r.id as string] || [];
          const act = (r.customer_id ? actBy[r.customer_id as string] : '') || '';
          slips.push(repairSlip(r, insp2, act));
          const mamoru = (r.qty_mamoru as number) || 0, other = (r.qty_other as number) || 0;
          const inspHtml = insp2.length
            ? insp2.map((i) => `#${i.scissor_number} ${esc(i.scissor_type || '가위')}${i.comment ? ` — ${esc(i.comment)}` : ''}`).join('<br>')
            : '';
          listRows.push({
            group: '복원수리', no: (r.as_id as string) || '', name: `${r.name}${act ? ` (${act})` : ''}`, contact: (r.phone as string) || '',
            itemsHtml: inspHtml, qty: `마모루 ${mamoru}${other > 0 ? `·타사 ${other}` : ''}자루`,
            memo: [(r.memo as string) || '', (r.admin_note as string) || ''].filter(Boolean).join(' / '),
          });
        });
      }

      // 2) 주문 슬립
      const ordIds = [...checked.order];
      if (ordIds.length) {
        const [{ data: ords }, { data: items }, { data: sers }] = await Promise.all([
          db.from('orders').select('*').in('id', ordIds),
          db.from('order_items').select('*').in('order_id', ordIds),
          db.from('product_serials').select('serial_number, product_id, order_id').in('order_id', ordIds),
        ]);
        ordIds.map((id) => (ords || []).find((o: Record<string, unknown>) => o.id === id)).filter(Boolean).forEach((o: Record<string, unknown>) => {
          const its = (items || []).filter((it: Record<string, unknown>) => it.order_id === o.id);
          const addr = [o.recipient_postcode ? `(${o.recipient_postcode})` : '', o.recipient_address, o.recipient_address_detail].filter(Boolean).join(' ');
          const mapped: { product_name: string; quantity: number; serialStr: string }[] = its.map((it: Record<string, unknown>) => ({
            product_name: it.product_name as string, quantity: (it.quantity as number) || 0,
            serialStr: (sers || []).filter((s: Record<string, unknown>) => s.order_id === o.id && s.product_id === it.product_id).map((s: Record<string, unknown>) => s.serial_number).join(', '),
          }));
          slips.push(saleSlip({
            headerLabel: 'MAMORU 출고 준비표', orderNo: (o.imweb_order_no as string) || '', tag: '주문',
            dateStr: (o.ordered_at as string) || '', custName: (o.recipient_name as string) || (o.orderer_name as string) || '', showNim: true,
            phone: (o.recipient_phone as string) || '', addr, items: mapped, memo: (o.recipient_memo as string) || '',
          }));
          listRows.push({
            group: '주문', no: (o.imweb_order_no as string) || '', name: (o.recipient_name as string) || (o.orderer_name as string) || '', contact: (o.recipient_phone as string) || '',
            itemsHtml: mapped.map((m) => `${esc(m.product_name)}${m.serialStr ? ` <span class="ser">(${esc(m.serialStr)})</span>` : ''} <b>×${m.quantity}</b>`).join('<br>'),
            qty: `${mapped.reduce((a, m) => a + m.quantity, 0)}개`, memo: (o.recipient_memo as string) || '',
          });
        });
      }

      // 3) 판매 슬립
      const saleIds = [...checked.sale];
      if (saleIds.length) {
        const [{ data: sales }, { data: items }, { data: sers }] = await Promise.all([
          db.from('offline_sales').select('*').in('id', saleIds),
          db.from('offline_sale_items').select('*').in('sale_id', saleIds),
          db.from('product_serials').select('serial_number, product_id, sale_item_id, offline_sale_id').in('offline_sale_id', saleIds),
        ]);
        const custIds = [...new Set((sales || []).map((s: Record<string, unknown>) => s.customer_id).filter(Boolean))] as string[];
        const addrBy: Record<string, string> = {};
        if (custIds.length) {
          const { data: custs } = await db.from('customers').select('id, postcode, address_road, address_detail').in('id', custIds);
          (custs || []).forEach((c: Record<string, unknown>) => { addrBy[c.id as string] = [c.postcode ? `(${c.postcode})` : '', c.address_road, c.address_detail].filter(Boolean).join(' '); });
        }
        saleIds.map((id) => (sales || []).find((s: Record<string, unknown>) => s.id === id)).filter(Boolean).forEach((s: Record<string, unknown>) => {
          const its = (items || []).filter((it: Record<string, unknown>) => it.sale_id === s.id);
          const isD = s.customer_type === 'dealer' || s.customer_type === 'academy';
          const mapped: { product_name: string; quantity: number; serialStr: string }[] = its.map((it: Record<string, unknown>) => ({
            product_name: it.product_name as string, quantity: (it.quantity as number) || 0,
            serialStr: (sers || []).filter((se: Record<string, unknown>) => se.offline_sale_id === s.id && (se.sale_item_id === it.id || se.product_id === it.product_id)).map((se: Record<string, unknown>) => se.serial_number).join(', '),
          }));
          slips.push(saleSlip({
            headerLabel: 'MAMORU 출고 준비표', orderNo: (s.sale_number as string) || '', tag: isD ? '거래처' : null,
            dateStr: (s.sale_date as string) || '', custName: (s.customer_name as string) || '', showNim: !isD,
            phone: (s.customer_phone as string) || '', addr: s.customer_id ? (addrBy[s.customer_id as string] || '') : '',
            items: mapped, memo: (s.memo as string) || '',
          }));
          listRows.push({
            group: '판매', no: (s.sale_number as string) || '', name: `${s.customer_name || ''}${isD ? ' (거래처)' : ''}`, contact: (s.customer_phone as string) || '',
            itemsHtml: mapped.map((m) => `${esc(m.product_name)}${m.serialStr ? ` <span class="ser">(${esc(m.serialStr)})</span>` : ''} <b>×${m.quantity}</b>`).join('<br>'),
            qty: `${mapped.reduce((a, m) => a + m.quantity, 0)}개`, memo: (s.memo as string) || '',
          });
        });
      }

      if (slips.length === 0) return;
      const w = window.open('', '_blank');
      if (!w) return;
      w.document.write(mode === 'tray' ? wrapTray(slips, '통합 준비표') : buildListDoc(listRows, '통합 준비표'));
      w.document.close();
      w.print();
    } finally {
      setPrinting(false);
    }
  };

  const active = rows[tab];
  const activeChecked = checked[tab];
  const counts = useMemo(() => ({
    repair: checked.repair.size, order: checked.order.size, sale: checked.sale.size,
  }), [checked]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl flex flex-col w-full" style={{ maxWidth: 620, maxHeight: '88vh' }}
        onClick={(e) => e.stopPropagation()}>
        {/* 헤더 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">통합 준비표 <span className="font-normal text-neutral-400">지금 준비할 것</span></h3>
          <div className="flex items-center gap-2">
            <div className="flex items-center text-[11px] rounded-md overflow-hidden border border-neutral-200">
              <button onClick={() => setMode('tray')} className={`px-2 py-1 font-semibold transition ${mode === 'tray' ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500 hover:bg-neutral-50'}`}>트레이형</button>
              <button onClick={() => setMode('list')} className={`px-2 py-1 font-semibold transition ${mode === 'list' ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500 hover:bg-neutral-50'}`}>리스트형</button>
            </div>
            <button onClick={handlePrint} disabled={printing || totalChecked === 0}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition disabled:opacity-40">
              <Printer size={12} /> {printing ? '준비 중…' : `인쇄 (${totalChecked})`}
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
          </div>
        </div>

        {/* 탭 */}
        <div className="flex gap-1 px-3 pt-2 border-b border-neutral-100">
          {TABS.map((t) => {
            const isA = tab === t.key;
            const n = rows[t.key].length;
            const c = counts[t.key];
            return (
              <button key={t.key} onClick={() => setTab(t.key)}
                className={`flex items-center gap-1.5 px-3 py-2 text-sm font-semibold border-b-2 transition ${isA ? 'border-stone-900 text-stone-900' : 'border-transparent text-stone-400 hover:text-stone-600'}`}>
                {t.label}
                <span className={`text-[11px] rounded-full px-1.5 ${c > 0 ? 'bg-stone-900 text-white' : 'bg-neutral-200 text-neutral-500'}`}>{c > 0 ? `${c}/${n}` : n}</span>
              </button>
            );
          })}
        </div>

        {/* 체크리스트 */}
        <div className="overflow-y-auto flex-1 p-3">
          {loading ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">불러오는 중…</div>
          ) : active.length === 0 ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">준비 대기 건이 없습니다</div>
          ) : (
            <>
              <button onClick={() => toggleAll(tab)} className="text-xs text-blue-600 hover:text-blue-700 mb-1.5 px-1">
                {activeChecked.size === active.length ? '전체 해제' : '전체 선택'}
              </button>
              <div className="space-y-1">
                {active.map((r) => {
                  const on = activeChecked.has(r.id);
                  return (
                    <label key={r.id} className={`flex items-center gap-2.5 px-2.5 py-2 rounded-lg border cursor-pointer transition ${on ? 'border-stone-900 bg-stone-50' : 'border-neutral-200 hover:bg-neutral-50'}`}>
                      <input type="checkbox" checked={on} onChange={() => toggle(tab, r.id)} className="w-4 h-4 accent-stone-900" />
                      <div className="min-w-0 flex-1">
                        <div className="flex items-center gap-1.5">
                          <span className="font-semibold text-sm text-indigo-black truncate">{r.name}</span>
                          {r.label && <span className="font-mono text-[11px] text-neutral-400 truncate">{r.label}</span>}
                        </div>
                        {r.sub && <div className="text-[11px] text-neutral-500">{r.sub}</div>}
                      </div>
                      <span className={`shrink-0 text-[10px] font-semibold px-1.5 py-0.5 rounded ${r.hasInvoice ? 'bg-emerald-50 text-emerald-700' : 'bg-neutral-100 text-neutral-400'}`}>
                        {r.hasInvoice ? '📦 송장' : '송장전'}
                      </span>
                    </label>
                  );
                })}
              </div>
            </>
          )}
        </div>

        {/* 푸터 */}
        <div className="px-4 py-2.5 border-t border-neutral-200 flex items-center justify-between text-xs text-neutral-500">
          <span>선택 <b className="text-stone-900">{totalChecked}</b>건 (복원 {counts.repair}·주문 {counts.order}·판매 {counts.sale})</span>
          <span className="text-neutral-400">체크는 인쇄 대상 선택일 뿐 데이터는 변하지 않습니다</span>
        </div>
      </div>
    </div>
  );
}
