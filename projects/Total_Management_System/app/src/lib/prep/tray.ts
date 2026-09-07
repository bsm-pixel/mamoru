/**
 * 준비표 트레이 공용 빌더 — 복원수리·판매·주문·납품 슬립을 한 A4(2건/장)로 합본 출력.
 * 기존 prep-sheet-modal / repair-prep-sheet-modal 의 트레이 골격을 통합(105×297mm 슬립, 중앙 절취선).
 * 통합 준비표(unified-prep-modal)가 도메인 혼합 인쇄에 사용. (2026-09-07)
 */

export function esc(s: unknown): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
const won = (n: number) => `${(n || 0).toLocaleString()}원`;

// ── 복원수리 슬립 ──
export interface PrepInspection { scissor_number: number; scissor_type?: string | null; comment?: string | null }
export interface PrepRepair {
  as_id?: string | null; name?: string | null; phone?: string | null;
  proceed_type?: string | null; postcode?: string | null; address?: string | null; address_detail?: string | null;
  qty_mamoru?: number | null; qty_other?: number | null;
  memo?: string | null; admin_note?: string | null;
  service_cost?: number | null; shipping_fee?: number | null; total_amount?: number | null;
  received_at?: string | null;
}
export function repairSlip(r: PrepRepair, insp: PrepInspection[], activity: string): string {
  const addr = [r.postcode ? `(${r.postcode})` : '', r.address, r.address_detail].filter(Boolean).join(' ');
  const mamoru = r.qty_mamoru || 0, other = r.qty_other || 0, total = mamoru + other;
  const inspByNo: Record<number, PrepInspection> = {};
  insp.forEach((i) => { inspByNo[i.scissor_number] = i; });
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
}

// ── 판매·주문·납품 슬립(공용) ──
export interface PrepSaleItem { product_name?: string | null; quantity?: number | null; serialStr?: string }
export interface PrepSale {
  headerLabel: string;          // "MAMORU 출고 준비표"
  orderNo?: string | null;      // 판매번호 / 주문번호
  tag?: string | null;          // "거래처" / "주문" 등 (없으면 미표시)
  dateStr?: string | null;
  custName?: string | null;
  showNim?: boolean;            // 고객 님 표기(거래처는 false)
  phone?: string | null;
  addr?: string | null;
  items: PrepSaleItem[];
  memo?: string | null;
}
export function saleSlip(s: PrepSale): string {
  const itemsHtml = s.items.map((it) =>
    `<div class="it">☐ <b>${esc(it.product_name)}</b>${it.serialStr ? ` <span class="ser">(${esc(it.serialStr)})</span>` : ''}<span class="q">×${it.quantity || 0}</span></div>`
  ).join('') || '<div class="it muted">품목 없음</div>';
  return `<div class="slip">
    <div class="hd">${esc(s.headerLabel)}</div>
    <div class="ono">${esc(s.orderNo)}${s.tag ? ` <span class="tag">${esc(s.tag)}</span>` : ''}</div>
    <div class="dt">${esc((s.dateStr || '').slice(0, 10))}</div>
    <div class="sec">
      <div class="cust">${esc(s.custName)}${s.showNim ? '<span class="nim"> 님</span>' : ''}</div>
      ${s.phone ? `<div class="ph">${esc(s.phone)}</div>` : ''}
      ${s.addr ? `<div class="addr">${esc(s.addr)}</div>` : ''}
    </div>
    <div class="sec items">${itemsHtml}</div>
    ${s.memo ? `<div class="sec memo"><span class="lbl">메모</span> ${esc(s.memo)}</div>` : ''}
    <div class="chk">☐ 포장 완료</div>
    <div class="ft">MAMORU</div>
  </div>`;
}

// ── 트레이 문서(2슬립/장) ──
export const TRAY_CSS = `
  @page { size: A4; margin: 0; }
  * { box-sizing: border-box; margin: 0; }
  body { font-family:'Noto Sans KR','Apple SD Gothic Neo',sans-serif; color:#000; }
  .page { width:210mm; height:297mm; display:flex; page-break-after:always; }
  .page:last-child { page-break-after:auto; }
  .slip { width:105mm; height:297mm; padding:11mm 8mm 9mm; display:flex; flex-direction:column;
          border-right:1px dashed #888; overflow:hidden; }
  .slip.empty, .page > .slip:last-child { border-right:0; }
  .hd { font-size:10px; letter-spacing:2px; color:#666; }
  .ono { font-family:'Courier New',monospace; font-weight:800; font-size:17px; margin-top:3mm; letter-spacing:.4px; }
  .tag { font-size:9px; background:#000; color:#fff; padding:1px 5px; border-radius:3px; font-family:sans-serif; vertical-align:middle; }
  .dt { font-size:11px; color:#888; margin-top:1mm; }
  .sec { border-top:1px solid #000; margin-top:4mm; padding-top:3mm; }
  .cust { font-size:23px; font-weight:800; line-height:1.15; }
  .cust .act { font-size:13px; font-weight:600; color:#444; }
  .nim { font-size:13px; font-weight:400; color:#666; }
  .ph { font-size:14px; margin-top:1.5mm; }
  .addr { font-size:12px; color:#333; margin-top:1.5mm; line-height:1.45; }
  .qty { font-size:14px; } .qty b { font-size:16px; } .qty .tot { color:#888; font-size:11px; }
  .items .it, .scissors .it { font-size:13px; padding:2mm 0; border-bottom:1px dotted #ddd; overflow:hidden; }
  .items .it:last-child, .scissors .it:last-child { border-bottom:0; }
  .it .ser { color:#888; font-size:10px; } .it .q { float:right; font-weight:700; }
  .scissors .no { color:#999; font-weight:700; margin-right:1mm; }
  .scissors .cm { color:#666; font-size:11px; } .scissors .blank { color:#ccc; letter-spacing:1px; }
  .it.muted { color:#aaa; }
  .cost { font-size:13px; }
  .memo { font-size:12px; line-height:1.55; white-space:pre-wrap; } .memo .lbl { font-weight:700; }
  .chk { margin-top:auto; border-top:2px solid #000; padding-top:3mm; font-size:15px; font-weight:600; }
  .ft { text-align:center; font-size:9px; color:#ccc; margin-top:3mm; letter-spacing:3px; }
  @media print { body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }`;

export function wrapTray(slips: string[], title = '준비표'): string {
  const pages: string[] = [];
  for (let i = 0; i < slips.length; i += 2) {
    const left = slips[i];
    const right = slips[i + 1] || '<div class="slip empty"></div>';
    pages.push(`<div class="page">${left}${right}</div>`);
  }
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${esc(title)}</title>
  <style>${TRAY_CSS}</style></head><body>${pages.join('')}</body></html>`;
}
