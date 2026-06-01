import { renderToStaticMarkup } from 'react-dom/server';
import QRCode from 'react-qr-code';

/**
 * 소싱 라벨 브라우저 인쇄 (공용) — STEP 1 품목별 칩 + STEP 2 전체인쇄가 공유.
 * 라벨 = QR + 번호(#001) + 품목명 + 한화가격. 매수만큼 반복.
 * QR은 react-qr-code 를 문자열로 렌더(renderToStaticMarkup) → DOM 의존 없음.
 * ※ 읽기 전용 — 어떤 데이터도 수정하지 않음.
 */

export interface PrintLabelInput {
  sticker_no: string;
  product_name: string;
  unit_price: number; // CNY
  qrValue: string; // QR에 인코딩할 URL
  copies: number; // 매수
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function printSourcingLabels(
  labels: PrintLabelInput[],
  size: { w: number; h: number }, // mm
  exchangeRate: number
) {
  if (labels.length === 0) return;

  // 폰트 스케일 (LabelPreview 와 동일 규칙: 40×20 짧은변 20mm = 1.0)
  const shortSide = Math.min(size.w, size.h);
  const fontScale = shortSide / 20;
  const numFontPt = Math.max(10, Math.min(18, 11 * fontScale));
  const nameFontPt = Math.max(9, Math.min(16, 10 * fontScale));
  const priceFontPt = Math.max(8, Math.min(14, 9 * fontScale));
  const qrMm = Math.min(size.w, size.h) * 0.8;
  const krwOf = (cny: number) => Math.round((cny || 0) * (exchangeRate || 0));

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
            <div class="price">₩${krwOf(lab.unit_price).toLocaleString()}</div>
          </div>
        </div>`;
      return Array.from({ length: Math.max(1, lab.copies) }, () => single).join('');
    })
    .join('');

  const win = window.open('', '_blank', 'width=720,height=560');
  if (!win) {
    alert('팝업 차단을 해제해 주세요.');
    return;
  }
  win.document.open();
  win.document.write(`<!DOCTYPE html>
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
    .name { font-size: ${nameFontPt}pt; margin-top: 0.6mm; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; word-break: break-all; line-height: 1.15; text-align: center; }
    .price { font-size: ${priceFontPt}pt; font-weight: 700; margin-top: 0.5mm; line-height: 1.1; }
  </style>
</head>
<body>
  ${labelHtml}
  <script>
    window.addEventListener('load', function() {
      setTimeout(function() { window.focus(); window.print(); }, 150);
    });
    window.addEventListener('afterprint', function() {
      setTimeout(function() { window.close(); }, 100);
    });
  </script>
</body>
</html>`);
  win.document.close();
}
