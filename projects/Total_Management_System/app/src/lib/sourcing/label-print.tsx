import { renderToStaticMarkup } from 'react-dom/server';
import QRCode from 'react-qr-code';

/**
 * 소싱 라벨/리스트 인쇄 (공용). ※ 읽기 전용 — 데이터 수정 없음.
 * - printSourcingLabels: 라벨 = QR + 번호(#001) + 품목명. (가격 미표시 — 사장님 요청)
 * - printSourcingPriceList: A4 가격 리스트 = 번호 · 품목명 · 단가(¥) · 가격(₩).
 */

export interface PrintLabelInput {
  sticker_no: string;
  product_name: string;
  qrValue: string; // QR에 인코딩할 URL
  copies: number;
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function openPrint(html: string) {
  const win = window.open('', '_blank', 'width=820,height=640');
  if (!win) {
    alert('팝업 차단을 해제해 주세요.');
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
}

const PRINT_SCRIPT = `<script>
  window.addEventListener('load', function() {
    setTimeout(function() { window.focus(); window.print(); }, 150);
  });
  window.addEventListener('afterprint', function() {
    setTimeout(function() { window.close(); }, 100);
  });
</script>`;

/** 라벨 인쇄 (QR + 번호 + 품목명, 가격 없음) */
export function printSourcingLabels(labels: PrintLabelInput[], size: { w: number; h: number }) {
  if (labels.length === 0) return;

  const shortSide = Math.min(size.w, size.h);
  const fontScale = shortSide / 20;
  const numFontPt = Math.max(11, Math.min(20, 12 * fontScale));
  const nameFontPt = Math.max(10, Math.min(18, 11 * fontScale));
  const qrMm = Math.min(size.w, size.h) * 0.8;

  const labelHtml = labels
    .map((lab) => {
      const qrSvg = renderToStaticMarkup(<QRCode value={lab.qrValue} size={256} level="M" />);
      const seq = lab.sticker_no.split('-').pop() || '000';
      const single = `
        <div class="label">
          <div class="qr">${qrSvg}</div>
          <div class="text">
            <div class="num">#${esc(seq)}</div>
            <div class="name">${esc(lab.product_name)}</div>
          </div>
        </div>`;
      return Array.from({ length: Math.max(1, lab.copies) }, () => single).join('');
    })
    .join('');

  openPrint(`<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title></title>
  <style>
    @page { size: ${size.w}mm ${size.h}mm; margin: 0; }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; background: #fff; font-family: 'Noto Sans KR', -apple-system, sans-serif; color: #000; }
    .label {
      width: ${size.w}mm; height: ${size.h}mm; padding: 1mm;
      display: grid; grid-template-columns: ${qrMm}mm ${size.w - qrMm - 2}mm; gap: 1mm;
      align-items: center; overflow: hidden; page-break-after: always; break-after: page;
    }
    .label:last-child { page-break-after: auto; break-after: auto; }
    .qr { width: ${qrMm}mm; height: ${qrMm}mm; display: flex; align-items: center; justify-content: center; }
    .qr svg { width: ${qrMm}mm !important; height: ${qrMm}mm !important; display: block; }
    .text { padding-left: 0.5mm; line-height: 1.15; overflow: hidden; display: flex; flex-direction: column; justify-content: center; text-align: center; align-items: center; }
    .num { font-size: ${numFontPt}pt; font-weight: 800; letter-spacing: 0.3px; line-height: 1.1; }
    .name { font-size: ${nameFontPt}pt; margin-top: 0.8mm; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; word-break: break-all; line-height: 1.15; text-align: center; }
  </style>
</head>
<body>${labelHtml}${PRINT_SCRIPT}</body>
</html>`);
}

export interface PriceListRow {
  seq: string;
  product_name: string;
  unit_price: number; // CNY
  supplier_name?: string | null;
}

/** A4 가격 리스트 인쇄 (번호 · 품목명 · 업체 · ¥단가 · ₩가격) */
export function printSourcingPriceList(rows: PriceListRow[], exchangeRate: number, title: string) {
  if (rows.length === 0) return;
  const krwOf = (c: number) => Math.round((c || 0) * (exchangeRate || 0));

  const body = rows
    .map(
      (r) => `<tr>
        <td class="c">${esc(r.seq)}</td>
        <td>${esc(r.product_name || '')}</td>
        <td class="muted">${esc(r.supplier_name || '')}</td>
        <td class="r">¥${r.unit_price}</td>
        <td class="r b">₩${krwOf(r.unit_price).toLocaleString()}</td>
      </tr>`
    )
    .join('');

  openPrint(`<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title></title>
  <style>
    @page { size: A4; margin: 14mm; }
    * { box-sizing: border-box; }
    body { margin: 0; font-family: 'Noto Sans KR', -apple-system, sans-serif; color: #111; }
    h1 { font-size: 16pt; margin: 0 0 2px; }
    .sub { font-size: 9pt; color: #666; margin-bottom: 12px; }
    table { width: 100%; border-collapse: collapse; font-size: 10.5pt; }
    th, td { border: 1px solid #ccc; padding: 6px 8px; text-align: left; }
    th { background: #f3f3f3; font-size: 9.5pt; }
    td.c { text-align: center; font-family: monospace; font-weight: 700; width: 48px; }
    td.r { text-align: right; white-space: nowrap; }
    td.b { font-weight: 700; }
    td.muted { color: #777; font-size: 9pt; }
    tr { page-break-inside: avoid; }
  </style>
</head>
<body>
  <h1>${esc(title)} — 품목 가격 리스트</h1>
  <div class="sub">총 ${rows.length}품목 · 환율 1¥ = ${exchangeRate}원</div>
  <table>
    <thead><tr><th>번호</th><th>품목명</th><th>업체</th><th>단가(¥)</th><th>가격(₩)</th></tr></thead>
    <tbody>${body}</tbody>
  </table>
  ${PRINT_SCRIPT}
</body>
</html>`);
}
