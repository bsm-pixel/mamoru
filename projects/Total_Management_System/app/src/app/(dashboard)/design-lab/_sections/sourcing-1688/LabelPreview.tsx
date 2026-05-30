'use client';

import { useMemo, useRef, useState } from 'react';
import QRCode from 'react-qr-code';
import { Printer, FileImage, AlertCircle, FileDown } from 'lucide-react';
import type { DemoPO, DemoPOItem } from './types';

/**
 * 50×30mm 기본 + 사이즈 선택 가능한 열전사 라벨 미리보기.
 *
 * 동작:
 * - 화면 미리보기: px 단위(2배 스케일)로 가독성 우선
 * - 인쇄: @media print 에서 mm 단위로 강제 override → 실제 라벨지 사이즈 정확
 * - 테스트 1장: data-mode="test" 시 첫 라벨만 보이도록 CSS 필터
 * - PDF 저장: 시스템 인쇄 다이얼로그에서 "PDF로 저장" 선택 (별도 코드 X)
 * - PNG: html2canvas (이미 설치됨)
 */

type LabelSize = {
  id: string;
  name: string;
  w: number; // mm
  h: number; // mm
};

const LABEL_PRESETS: LabelSize[] = [
  { id: 'p30x15', name: '3 × 1.5 cm (30 × 15 mm)', w: 30, h: 15 },
  { id: 'p40x20', name: '4 × 2 cm (40 × 20 mm)', w: 40, h: 20 },
  { id: 'p40x30', name: '4 × 3 cm (40 × 30 mm)', w: 40, h: 30 },
];

const MM_TO_PX = 3.78; // 96dpi 기준
const SCREEN_SCALE = 2; // 화면 미리보기는 2배 키워서 보기
const DEMO_BASE_URL = 'https://app-eta-sandy-75.vercel.app/purchasing/inbound';

export function LabelPreview({ po }: { po: DemoPO }) {
  const items = useMemo(() => po.items.filter((it) => it.product_name), [po.items]);
  const [sizeId, setSizeId] = useState<string>('p40x20');
  const [customW, setCustomW] = useState(40);
  const [customH, setCustomH] = useState(20);
  const [printMode, setPrintMode] = useState<'idle' | 'test' | 'all'>('idle');
  const [pngBusy, setPngBusy] = useState(false);
  const printAreaRef = useRef<HTMLDivElement>(null);

  const size = useMemo<LabelSize>(() => {
    if (sizeId === 'custom') {
      const w = Math.max(15, Math.min(200, customW || 50));
      const h = Math.max(10, Math.min(200, customH || 30));
      return { id: 'custom', name: `${w} × ${h}`, w, h };
    }
    return LABEL_PRESETS.find((p) => p.id === sizeId) || LABEL_PRESETS[1];
  }, [sizeId, customW, customH]);

  const handlePrint = (mode: 'test' | 'all') => {
    if (!printAreaRef.current) return;
    const cardsInDom = Array.from(
      printAreaRef.current.querySelectorAll('.label-card')
    ) as HTMLElement[];
    if (cardsInDom.length === 0) return;

    const targetItems = mode === 'test' ? items.slice(0, 1) : items;
    const cards = mode === 'test' ? cardsInDom.slice(0, 1) : cardsInDom;

    // 각 라벨의 QR SVG outerHTML 추출 (라이브 DOM 그대로 사용)
    const labelHtml = targetItems
      .map((it, i) => {
        const card = cards[i];
        const svgEl = card?.querySelector('svg');
        const qrSvg = svgEl ? svgEl.outerHTML : '';
        const seq = it.sticker_no.split('-').pop() || '000';
        const poShort = po.po_number.replace('PO-', '').replace('-DEMO', '');
        const moqText = it.moq ? ` · MOQ ${it.moq}` : '';
        return `
          <div class="label">
            <div class="qr">${qrSvg}</div>
            <div class="text">
              <div class="num">#${escapeHtml(seq)}</div>
              <div class="name">${escapeHtml(it.product_name)}</div>
              <div class="price">¥${it.unit_price} × ${it.quantity}${escapeHtml(moqText)}</div>
              <div class="po">${escapeHtml(poShort)}</div>
            </div>
          </div>
        `;
      })
      .join('');

    // 새 창에 인쇄 전용 HTML 주입 — 부모 페이지의 빈 영역이 페이지 수를 늘리지 않음
    const win = window.open('', '_blank', 'width=720,height=560');
    if (!win) {
      alert('팝업 차단을 해제해 주세요.');
      return;
    }
    win.document.open();
    // 머리글·바닥글 제거: title 비우기 (브라우저 자동 머리글 = 페이지 title)
    win.document.write(`<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <title></title>
  <style>
    @page {
      size: ${size.w}mm ${size.h}mm;
      margin: 0;
    }
    * { box-sizing: border-box; }
    html, body {
      margin: 0;
      padding: 0;
      background: #fff;
      font-family: 'Noto Sans KR', -apple-system, sans-serif;
      color: #000;
    }
    .label {
      width: ${size.w}mm;
      height: ${size.h}mm;
      padding: 1mm;
      display: grid;
      grid-template-columns: ${qrMm}mm ${size.w - qrMm - 2}mm;
      gap: 1mm;
      align-items: center;
      overflow: hidden;
      page-break-after: always;
      break-after: page;
    }
    .label:last-child {
      page-break-after: auto;
      break-after: auto;
    }
    .qr {
      width: ${qrMm}mm;
      height: ${qrMm}mm;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    .qr svg {
      width: ${qrMm}mm !important;
      height: ${qrMm}mm !important;
      display: block;
    }
    .text {
      padding-left: 0.5mm;
      line-height: 1.15;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      justify-content: center;
      text-align: center;
      align-items: center;
    }
    .num {
      font-size: ${numFontPt}pt;
      font-weight: 800;
      letter-spacing: 0.3px;
      line-height: 1.1;
    }
    .name {
      font-size: ${nameFontPt}pt;
      margin-top: 0.6mm;
      display: -webkit-box;
      -webkit-line-clamp: 2;
      -webkit-box-orient: vertical;
      overflow: hidden;
      word-break: break-all;
      line-height: 1.15;
      text-align: center;
    }
    .price {
      font-size: ${priceFontPt}pt;
      color: #444;
      margin-top: 0.5mm;
    }
    .po {
      font-size: ${poFontPt}pt;
      color: #888;
      font-family: monospace;
      margin-top: 0.4mm;
    }
  </style>
</head>
<body>
  ${labelHtml}
  <script>
    window.addEventListener('load', function() {
      setTimeout(function() {
        window.focus();
        window.print();
      }, 150);
    });
    window.addEventListener('afterprint', function() {
      setTimeout(function() { window.close(); }, 100);
    });
  </script>
</body>
</html>`);
    win.document.close();
    setPrintMode(mode);
    setTimeout(() => setPrintMode('idle'), 200);
  };

  const handleCsv = () => {
    if (items.length === 0) return;
    // UTF-8 BOM + CRLF (Excel/NiceLabel 호환)
    const headers = [
      'sticker_no',
      'product_name',
      'qr_url',
      'unit_price_cny',
      'quantity',
      'moq',
      'po_number',
      'vendor_url',
      'features_memo',
    ];
    const esc = (v: string | number | null | undefined): string => {
      const s = v == null ? '' : String(v);
      return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = [
      headers.join(','),
      ...items.map((it) =>
        [
          it.sticker_no,
          it.product_name,
          `${DEMO_BASE_URL}/${it.id}`,
          it.unit_price,
          it.quantity,
          it.moq ?? '',
          po.po_number,
          it.vendor_url,
          it.features_memo,
        ]
          .map(esc)
          .join(',')
      ),
    ];
    const csv = '﻿' + lines.join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `labels-${po.po_number}-${items.length}items.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  };

  const handlePng = async () => {
    if (!printAreaRef.current) return;
    setPngBusy(true);
    try {
      const { default: html2canvas } = await import('html2canvas');
      const canvas = await html2canvas(printAreaRef.current, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
      });
      canvas.toBlob((blob) => {
        if (!blob) return;
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `labels-${po.po_number}-${size.w}x${size.h}mm.png`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      });
    } catch (e) {
      // 데모이므로 silent fail
      console.error('PNG export failed', e);
    } finally {
      setPngBusy(false);
    }
  };

  if (items.length === 0) {
    return (
      <div className="rounded-xl border border-dashed border-stone-300 bg-stone-50 p-8 text-center">
        <Printer className="mx-auto text-stone-300 mb-2" size={28} />
        <p className="text-xs text-stone-500">
          품목명을 입력하면 라벨 미리보기가 여기에 표시됩니다.
        </p>
      </div>
    );
  }

  // 화면 미리보기 px
  const labelW = size.w * MM_TO_PX * SCREEN_SCALE;
  const labelH = size.h * MM_TO_PX * SCREEN_SCALE;
  // QR은 짧은 변의 80% (양쪽 여백 10%)
  const qrMm = Math.min(size.w, size.h) * 0.8;
  const qrPx = qrMm * MM_TO_PX * SCREEN_SCALE;
  // 텍스트 영역
  const textColMm = size.w - qrMm - 2; // padding 1mm × 2
  const textColPx = textColMm * MM_TO_PX * SCREEN_SCALE;

  // 사이즈에 따른 폰트 스케일 (작은 라벨 = 작은 폰트)
  const shortSide = Math.min(size.w, size.h);
  const fontScale = shortSide / 30; // 30mm 짧은 변 기준
  const numFontPt = Math.max(8, Math.min(14, 10 * fontScale));
  const nameFontPt = Math.max(5, Math.min(9, 6.5 * fontScale));
  const priceFontPt = Math.max(4.5, Math.min(7, 5.5 * fontScale));
  const poFontPt = Math.max(4, Math.min(6, 4.5 * fontScale));

  return (
    <div className="space-y-3">
      {/* 컨트롤 바 */}
      <div className="label-controls flex items-end flex-wrap gap-3 p-3.5 bg-stone-900 text-white rounded-xl">
        <div>
          <label className="text-[10px] uppercase tracking-wider opacity-60 block mb-1 font-medium">
            라벨 사이즈
          </label>
          <select
            value={sizeId}
            onChange={(e) => setSizeId(e.target.value)}
            className="px-2.5 py-1.5 rounded-lg bg-white/10 text-white text-xs font-medium border border-white/15 focus:outline-none focus:ring-2 focus:ring-white/30"
          >
            {LABEL_PRESETS.map((p) => (
              <option key={p.id} value={p.id} className="text-stone-900">
                {p.name} mm
              </option>
            ))}
            <option value="custom" className="text-stone-900">
              사용자 지정
            </option>
          </select>
        </div>

        {sizeId === 'custom' && (
          <div className="flex items-end gap-1.5">
            <div>
              <label className="text-[10px] uppercase tracking-wider opacity-60 block mb-1 font-medium">
                가로 (mm)
              </label>
              <input
                type="number"
                min={15}
                max={200}
                value={customW}
                onChange={(e) => setCustomW(Number(e.target.value) || 50)}
                className="w-16 px-2 py-1.5 rounded-lg bg-white/10 text-white text-xs font-medium border border-white/15 focus:outline-none focus:ring-2 focus:ring-white/30"
              />
            </div>
            <div className="pb-1.5 text-white/40 text-sm">×</div>
            <div>
              <label className="text-[10px] uppercase tracking-wider opacity-60 block mb-1 font-medium">
                세로 (mm)
              </label>
              <input
                type="number"
                min={10}
                max={200}
                value={customH}
                onChange={(e) => setCustomH(Number(e.target.value) || 30)}
                className="w-16 px-2 py-1.5 rounded-lg bg-white/10 text-white text-xs font-medium border border-white/15 focus:outline-none focus:ring-2 focus:ring-white/30"
              />
            </div>
          </div>
        )}

        <div className="flex-1 min-w-[100px]" />

        <div className="flex items-end gap-2 flex-wrap">
          <button
            type="button"
            onClick={handleCsv}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-amber-400 text-stone-900 text-xs font-bold hover:bg-amber-300"
            title="NiceLabel에서 데이터 소스로 연결해 자동 N장 인쇄 (가장 정확)"
          >
            <FileDown size={13} /> NiceLabel용 CSV
          </button>
          <button
            type="button"
            onClick={() => handlePrint('test')}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/15 text-white text-xs font-medium hover:bg-white/25 border border-white/10"
            title="첫 번째 라벨 1장만 브라우저 인쇄 — 사이즈/여백 점검용"
          >
            <Printer size={13} /> 테스트 1장
          </button>
          <button
            type="button"
            onClick={() => handlePrint('all')}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white text-stone-900 text-xs font-bold hover:bg-stone-100"
            title="전체 인쇄 — 다이얼로그에서 'PDF로 저장' 선택 시 PDF로 저장"
          >
            <Printer size={13} /> 전체 {items.length}장 인쇄 / PDF
          </button>
          <button
            type="button"
            onClick={handlePng}
            disabled={pngBusy}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/15 text-white text-xs font-medium hover:bg-white/25 border border-white/10 disabled:opacity-50"
            title="현재 미리보기를 PNG 이미지로 저장"
          >
            <FileImage size={13} /> {pngBusy ? '생성 중…' : 'PNG'}
          </button>
        </div>
      </div>

      {/* 추천 경로 — NiceLabel CSV */}
      <div className="label-meta flex items-start gap-2 px-3 py-2.5 rounded-lg bg-amber-50 border-2 border-amber-300 text-[11px] text-stone-800 leading-relaxed">
        <FileDown size={14} className="text-amber-700 flex-shrink-0 mt-0.5" />
        <div>
          <strong className="text-stone-900">✨ 권장 워크플로 — NiceLabel + CSV (가장 정확)</strong>
          <ol className="mt-1 ml-4 space-y-0.5 list-decimal">
            <li>위 <strong>[NiceLabel용 CSV]</strong> 버튼 → CSV 다운로드</li>
            <li>NiceLabel에서 라벨 템플릿 1번 디자인 (QR / 번호 / 품목명 자리)</li>
            <li>CSV를 <strong>데이터 소스</strong>로 연결 → 자동 N장 인쇄 (위치·회전 NiceLabel이 보정)</li>
          </ol>
          <div className="mt-1.5 text-[10px] text-stone-500">
            CSV 컬럼: sticker_no, product_name, qr_url, unit_price_cny, quantity, moq, po_number, vendor_url, features_memo
          </div>
        </div>
      </div>

      {/* 브라우저 인쇄 가이드 (보조) */}
      <details className="label-meta">
        <summary className="cursor-pointer text-[11px] text-stone-500 hover:text-stone-700 select-none">
          ⋯ 브라우저 인쇄 ([테스트 1장] / [전체 인쇄] 사용 시 설정) 펼치기
        </summary>
        <div className="mt-2 flex items-start gap-2 px-3 py-2 rounded-lg bg-stone-50 border border-stone-200 text-[11px] text-stone-700 leading-relaxed">
          <AlertCircle size={13} className="text-stone-500 flex-shrink-0 mt-0.5" />
          <div>
            <strong>인쇄 다이얼로그 → &lsquo;설정 더보기&rsquo; 펼치고 ↓</strong>
            <ul className="mt-1 ml-3 space-y-0.5 list-disc">
              <li><strong>용지 크기</strong>: &lsquo;사용자 지정&rsquo; → <strong>{size.w} × {size.h} mm</strong></li>
              <li><strong>여백</strong>: 없음 / <strong>배율</strong>: 100%</li>
              <li><strong>머리글·바닥글</strong>: <span className="text-rose-600 font-bold">반드시 끄기</span></li>
              <li><strong>배경 그래픽</strong>: 켜기 (QR 출력)</li>
              <li>※ Rongta 같은 영수증 프린터 계열은 좌측 치우침 발생 가능 — 그때는 위 NiceLabel CSV 경로 사용</li>
            </ul>
          </div>
        </div>
      </details>

      <div className="label-meta text-xs text-stone-500">
        실제 출력 사이즈: <strong className="text-stone-900">{size.w} × {size.h} mm</strong> · 한 품목당 1매 · 총{' '}
        <strong className="text-stone-900">{items.length}장</strong>
      </div>

      {/* 라벨 미리보기 영역 (인쇄 타겟) */}
      <div
        ref={printAreaRef}
        className="label-print-area"
        data-mode={printMode}
      >
        <div className="label-grid flex flex-wrap gap-4">
          {items.map((it) => (
            <div key={it.id} className="inline-flex flex-col items-start">
              <div className="label-stickerno text-[10px] text-stone-400 mb-1 font-mono">
                {it.sticker_no}
              </div>
              <LabelCard
                item={it}
                poNumber={po.po_number}
                labelW={labelW}
                labelH={labelH}
                qrPx={qrPx}
                textColPx={textColPx}
                numFontPt={numFontPt}
                nameFontPt={nameFontPt}
                priceFontPt={priceFontPt}
                poFontPt={poFontPt}
              />
            </div>
          ))}
        </div>
      </div>

      <div className="label-meta text-[10px] text-stone-400 italic">
        ※ 미리보기는 가독성을 위해 실제 크기의 2배로 표시됩니다. 실제 출력은 정확히 {size.w}×{size.h}mm.
      </div>
    </div>
  );
}

function LabelCard({
  item,
  poNumber,
  labelW,
  labelH,
  qrPx,
  textColPx,
  numFontPt,
  nameFontPt,
  priceFontPt,
  poFontPt,
}: {
  item: DemoPOItem;
  poNumber: string;
  labelW: number;
  labelH: number;
  qrPx: number;
  textColPx: number;
  numFontPt: number;
  nameFontPt: number;
  priceFontPt: number;
  poFontPt: number;
}) {
  const qrUrl = `${DEMO_BASE_URL}/${item.id}`;
  const seq = item.sticker_no.split('-').pop() || '000';
  // pt → px (화면용): 1pt ≈ 1.333px × SCREEN_SCALE
  const ptToPx = (pt: number) => pt * 1.333 * SCREEN_SCALE;

  return (
    <div
      className="label-card bg-white border border-stone-300 shadow-sm"
      style={{
        width: labelW,
        height: labelH,
        padding: 4 * SCREEN_SCALE,
        display: 'grid',
        gridTemplateColumns: `${qrPx}px ${textColPx}px`,
        gap: 4 * SCREEN_SCALE,
        alignItems: 'center',
        overflow: 'hidden',
      }}
    >
      <div className="qr-box flex items-center justify-center" style={{ width: qrPx, height: qrPx }}>
        <QRCode value={qrUrl} size={qrPx} level="M" />
      </div>
      <div
        className="label-text leading-tight"
        style={{ color: '#0a0a0a', overflow: 'hidden' }}
      >
        <div
          className="ln-num"
          style={{
            fontSize: ptToPx(numFontPt),
            fontWeight: 800,
            letterSpacing: 0.3,
            lineHeight: 1.1,
          }}
        >
          #{seq}
        </div>
        <div
          className="ln-name"
          style={{
            fontSize: ptToPx(nameFontPt),
            marginTop: 2 * SCREEN_SCALE,
            display: '-webkit-box',
            WebkitLineClamp: 2,
            WebkitBoxOrient: 'vertical',
            overflow: 'hidden',
            wordBreak: 'break-all',
            lineHeight: 1.15,
          }}
        >
          {item.product_name}
        </div>
        <div
          className="ln-price"
          style={{
            fontSize: ptToPx(priceFontPt),
            marginTop: 2 * SCREEN_SCALE,
            color: '#444',
          }}
        >
          ¥{item.unit_price} × {item.quantity}
          {item.moq ? ` · MOQ ${item.moq}` : ''}
        </div>
        <div
          className="ln-po"
          style={{
            fontSize: ptToPx(poFontPt),
            marginTop: 2 * SCREEN_SCALE,
            color: '#888',
            fontFamily: 'monospace',
          }}
        >
          {poNumber.replace('PO-', '').replace('-DEMO', '')}
        </div>
      </div>
    </div>
  );
}

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
