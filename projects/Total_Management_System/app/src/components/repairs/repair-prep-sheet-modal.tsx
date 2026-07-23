'use client';

import { useState, useEffect, useMemo } from 'react';
import { Printer } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';

/**
 * 복원수리 준비표 (트레이형) — 2026-07-23
 * A4 1장에 접수 2건(좌·우 105×297mm), 가운데 세로 절취선. 단건(상세)이면 왼쪽 1열만.
 * 판매 준비표(prep-sheet-modal)와 트레이 HTML 골격만 공유, 데이터는 repairs 전용.
 */

interface RepairRow {
  id: string; as_id: string; name: string; phone?: string | null; customer_id?: string | null;
  proceed_type?: string | null; postcode?: string | null; address?: string | null; address_detail?: string | null;
  qty_mamoru?: number | null; qty_other?: number | null;
  memo?: string | null; admin_note?: string | null;
  service_cost?: number | null; shipping_fee?: number | null; total_amount?: number | null;
  received_at?: string | null;
  [key: string]: unknown;
}
interface Inspection { repair_id: string; scissor_number: number; scissor_type?: string | null; comment?: string | null }

interface Props {
  repairIds: string[];   // repairs.id 배열 (단건 상세면 [id])
  onClose: () => void;
}

function esc(s: unknown): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
const won = (n: number) => `${(n || 0).toLocaleString()}원`;

export function RepairPrepSheetModal({ repairIds, onClose }: Props) {
  const [repairs, setRepairs] = useState<RepairRow[]>([]);
  const [inspByRepair, setInspByRepair] = useState<Record<string, Inspection[]>>({});
  const [activities, setActivities] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (repairIds.length === 0) return;
    (async () => {
      setLoading(true);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = createClient() as any;
      const [repRes, inspRes] = await Promise.all([
        db.from('repairs').select('*').in('id', repairIds),
        db.from('repair_inspections').select('repair_id, scissor_number, scissor_type, comment').in('repair_id', repairIds).order('scissor_number'),
      ]);
      const reps: RepairRow[] = (repRes.data || []) as RepairRow[];
      // 선택 순서 유지
      reps.sort((a, b) => repairIds.indexOf(a.id) - repairIds.indexOf(b.id));
      setRepairs(reps);

      const map: Record<string, Inspection[]> = {};
      (inspRes.data || []).forEach((i: Inspection) => { (map[i.repair_id] ||= []).push(i); });
      setInspByRepair(map);

      const custIds = [...new Set(reps.map((r) => r.customer_id).filter(Boolean))] as string[];
      if (custIds.length > 0) {
        const { data: custs } = await db.from('customers').select('id, activity_name').in('id', custIds);
        const am: Record<string, string> = {};
        (custs || []).forEach((c: { id: string; activity_name?: string | null }) => { if (c.activity_name) am[c.id] = c.activity_name; });
        setActivities(am);
      }
      setLoading(false);
    })();
  }, [repairIds]);

  const buildSlip = (r: RepairRow): string => {
    const activity = r.customer_id ? activities[r.customer_id] : '';
    const addr = [r.postcode ? `(${r.postcode})` : '', r.address, r.address_detail].filter(Boolean).join(' ');
    const mamoru = r.qty_mamoru || 0, other = r.qty_other || 0, total = mamoru + other;
    const insp = inspByRepair[r.id] || [];
    const inspByNo: Record<number, Inspection> = {};
    insp.forEach((i) => { inspByNo[i.scissor_number] = i; });

    // 자루별 라인 — 검수 있으면 채우고, 없으면 손으로 적을 빈칸
    const rowN = Math.min(Math.max(total, insp.length), 20);
    const lines = rowN > 0
      ? Array.from({ length: rowN }).map((_, k) => {
          const no = k + 1; const i = inspByNo[no];
          const label = i ? `<b>${esc(i.scissor_type || '가위')}</b>${i.comment ? ` <span class="cm">— ${esc(i.comment)}</span>` : ''}`
            : '<span class="blank">________________________</span>';
          return `<div class="it">☐ <span class="no">#${no}</span> ${label}</div>`;
        }).join('')
      : '<div class="it muted">자루 수 미입력</div>';

    const cost = (r.total_amount || 0) > 0
      ? `<div class="sec cost">수리비 ${won(r.service_cost || 0)}${(r.shipping_fee || 0) > 0 ? ` · 수거비 ${won(r.shipping_fee || 0)}` : ''} · <b>합계 ${won(r.total_amount || 0)}</b></div>`
      : '';
    const memo = [
      r.memo ? `<div><span class="lbl">메모</span> ${esc(r.memo)}</div>` : '',
      r.admin_note ? `<div><span class="lbl">전달</span> ${esc(r.admin_note)}</div>` : '',
    ].filter(Boolean).join('');

    return `<div class="slip">
      <div class="hd">MAMORU 복원수리 준비표</div>
      <div class="ono">${esc(r.as_id)}${r.proceed_type ? ` <span class="tag">${esc(r.proceed_type)}</span>` : ''}</div>
      <div class="dt">${esc((r.received_at || '').slice(0, 10))} 접수</div>
      <div class="sec">
        <div class="cust">${esc(r.name)}${activity ? ` <span class="act">(${esc(activity)})</span>` : ''}<span class="nim"> 님</span></div>
        ${r.phone ? `<div class="ph">${esc(r.phone)}</div>` : ''}
        ${addr ? `<div class="addr">${esc(addr)}</div>` : ''}
      </div>
      <div class="sec qty">마모루 <b>${mamoru}</b>자루${other > 0 ? ` · 타사 <b>${other}</b>자루` : ''} <span class="tot">(총 ${total}자루)</span></div>
      <div class="sec scissors">${lines}</div>
      ${cost}
      ${memo ? `<div class="sec memo">${memo}</div>` : ''}
      <div class="chk">☐ 검수 완료&nbsp;&nbsp;&nbsp;☐ 포장 완료</div>
      <div class="ft">MAMORU</div>
    </div>`;
  };

  const buildTrayHtml = (): string => {
    const pages: string[] = [];
    for (let i = 0; i < repairs.length; i += 2) {
      const left = buildSlip(repairs[i]);
      const right = repairs[i + 1] ? buildSlip(repairs[i + 1]) : '<div class="slip empty"></div>';
      pages.push(`<div class="page">${left}${right}</div>`);
    }
    return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>복원수리 준비표</title>
    <style>
      @page { size: A4; margin: 0; }
      * { box-sizing: border-box; margin: 0; }
      body { font-family:'Noto Sans KR','Apple SD Gothic Neo',sans-serif; color:#000; }
      .page { width:210mm; height:297mm; display:flex; page-break-after:always; }
      .page:last-child { page-break-after:auto; }
      .slip { width:105mm; height:297mm; padding:11mm 8mm 9mm; display:flex; flex-direction:column;
              border-right:1px dashed #888; overflow:hidden; }
      .slip.empty, .page > .slip:last-child { border-right:0; }
      .hd { font-size:10px; letter-spacing:2px; color:#666; }
      .ono { font-family:'Courier New',monospace; font-weight:800; font-size:16px; margin-top:3mm; letter-spacing:.3px; }
      .tag { font-size:9px; background:#000; color:#fff; padding:1px 5px; border-radius:3px; font-family:sans-serif; vertical-align:middle; }
      .dt { font-size:11px; color:#888; margin-top:1mm; }
      .sec { border-top:1px solid #000; margin-top:4mm; padding-top:3mm; }
      .cust { font-size:22px; font-weight:800; line-height:1.15; }
      .cust .act { font-size:13px; font-weight:600; color:#444; }
      .nim { font-size:13px; font-weight:400; color:#666; }
      .ph { font-size:14px; margin-top:1.5mm; }
      .addr { font-size:12px; color:#333; margin-top:1.5mm; line-height:1.45; }
      .qty { font-size:14px; } .qty b { font-size:16px; } .qty .tot { color:#888; font-size:11px; }
      .scissors .it { font-size:13px; padding:2.2mm 0; border-bottom:1px dotted #ddd; overflow:hidden; }
      .scissors .it:last-child { border-bottom:0; }
      .scissors .no { color:#999; font-weight:700; margin-right:1mm; }
      .scissors .cm { color:#666; font-size:11px; }
      .scissors .blank { color:#ccc; letter-spacing:1px; }
      .it.muted { color:#aaa; }
      .cost { font-size:13px; }
      .memo { font-size:12px; line-height:1.55; white-space:pre-wrap; }
      .memo .lbl { font-weight:700; }
      .chk { margin-top:auto; border-top:2px solid #000; padding-top:3mm; font-size:15px; font-weight:600; }
      .ft { text-align:center; font-size:9px; color:#ccc; margin-top:3mm; letter-spacing:3px; }
      @media print { body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }
    </style></head><body>${pages.join('')}</body></html>`;
  };

  const trayHtml = useMemo(() => (repairs.length > 0 ? buildTrayHtml() : ''),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [repairs, inspByRepair, activities]);

  const handlePrint = () => {
    const w = window.open('', '_blank');
    if (!w) return;
    w.document.write(trayHtml || buildTrayHtml());
    w.document.close();
    w.print();
  };

  const previewW = 300;
  const scale = previewW / (210 * 3.7795);
  const pageCount = Math.ceil(repairs.length / 2);
  const previewH = Math.min(297 * 3.7795 * scale * pageCount, 460);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl flex flex-col" style={{ width: '760px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">복원수리 준비표 {repairs.length > 0 && `(${repairs.length}건)`}</h3>
          <div className="flex items-center gap-2">
            <button onClick={handlePrint} disabled={loading || repairs.length === 0}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition disabled:opacity-50">
              <Printer size={12} /> 인쇄
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
          </div>
        </div>

        <div className="overflow-y-auto flex-1 p-5">
          {loading ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">데이터 로딩 중...</div>
          ) : repairs.length === 0 ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">선택된 건이 없습니다</div>
          ) : (
            <div>
              <p className="text-xs text-neutral-500 mb-2">
                A4 1장에 <b>2건</b> · 가운데 <b>점선에서 세로로 절취</b> → 반쪽(105×297mm) 트레이용 · 총 <b>{pageCount}장</b>
                {repairs.length % 2 === 1 && <span className="text-neutral-400"> (마지막 장은 왼쪽만)</span>}
              </p>
              <div className="border border-neutral-200 rounded-lg overflow-hidden bg-neutral-100"
                style={{ width: previewW, height: previewH, overflowY: 'auto' }}>
                <div style={{ width: 210 * 3.7795, transform: `scale(${scale})`, transformOrigin: 'top left' }}>
                  <iframe title="복원수리 준비표 미리보기" srcDoc={trayHtml} scrolling="no"
                    style={{ width: 210 * 3.7795, height: 297 * 3.7795 * pageCount, border: 0, display: 'block' }} />
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
