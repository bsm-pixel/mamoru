'use client';

import { useMemo, useRef, useState } from 'react';
import QRCode from 'react-qr-code';
import { Printer, FileImage, AlertCircle, FileDown, FileCode } from 'lucide-react';
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
  const [zplBusy, setZplBusy] = useState(false);
  const [dpi, setDpi] = useState<203 | 300>(300); // ZPL 출력 해상도 (ZD421T=300, 일부 모델=203)
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

  // Zebra 네이티브 ZPL(.zpl) 저장
  // 한글 폰트 깨짐 방지를 위해 라벨 전체를 1비트 비트맵(^GFA)으로 래스터화 →
  // QR·한글·레이아웃이 미리보기와 100% 동일, 프린터 폰트 설치 불필요.
  const handleZpl = async () => {
    if (!printAreaRef.current || items.length === 0) return;
    setZplBusy(true);
    try {
      const { default: html2canvas } = await import('html2canvas');
      const dotsPerMm = dpi / 25.4;
      const pw = Math.round(size.w * dotsPerMm); // print width (dots)
      const ll = Math.round(size.h * dotsPerMm); // label length (dots)
      const bytesPerRow = Math.ceil(pw / 8);

      const cards = Array.from(
        printAreaRef.current.querySelectorAll('.label-card')
      ) as HTMLElement[];

      const blocks: string[] = [];
      for (let i = 0; i < items.length; i++) {
        const card = cards[i];
        if (!card) continue;

        // 고해상 렌더 후 타겟 dot 크기로 다운스케일 (QR 모듈 가독 위해 scale 3)
        const src = await html2canvas(card, {
          backgroundColor: '#ffffff',
          scale: 3,
          useCORS: true,
        });
        const dest = document.createElement('canvas');
        dest.width = pw;
        dest.height = ll;
        const ctx = dest.getContext('2d');
        if (!ctx) continue;
        ctx.imageSmoothingEnabled = false; // QR 엣지 뭉개짐 방지
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, pw, ll);
        ctx.drawImage(src, 0, 0, pw, ll);
        const img = ctx.getImageData(0, 0, pw, ll).data;

        // 1비트 패킹 (1=흑=출력), MSB=좌측 픽셀
        const bytes: number[] = [];
        for (let y = 0; y < ll; y++) {
          for (let b = 0; b < bytesPerRow; b++) {
            let byte = 0;
            for (let bit = 0; bit < 8; bit++) {
              const x = b * 8 + bit;
              let dark = 0;
              if (x < pw) {
                const idx = (y * pw + x) * 4;
                const lum =
                  0.299 * img[idx] + 0.587 * img[idx + 1] + 0.114 * img[idx + 2];
                dark = lum < 128 ? 1 : 0;
              }
              byte = (byte << 1) | dark;
            }
            bytes.push(byte);
          }
        }
        const total = bytes.length;
        const hex = bytes
          .map((v) => v.toString(16).padStart(2, '0'))
          .join('')
          .toUpperCase();

        // ^GFA,총바이트,총바이트,행당바이트,16진데이터
        blocks.push(
          `^XA\n^PW${pw}\n^LL${ll}\n^LH0,0\n^FO0,0^GFA,${total},${total},${bytesPerRow},${hex}^FS\n^PQ1\n^XZ`
        );
      }

      const zpl = blocks.join('\n');
      const blob = new Blob([zpl], { type: 'text/plain;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `labels-${po.po_number}-${items.length}items-${dpi}dpi.zpl`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      console.error('ZPL export failed', e);
      alert('ZPL 생성 실패 — 콘솔을 확인해 주세요.');
    } finally {
      setZplBusy(false);
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

        {/* ZPL 출력 해상도 — 프린터 모델에 맞춤 (ZD421T=300) */}
        <div>
          <label className="text-[10px] uppercase tracking-wider opacity-60 block mb-1 font-medium">
            ZPL 해상도
          </label>
          <select
            value={dpi}
            onChange={(e) => setDpi(Number(e.target.value) as 203 | 300)}
            className="px-2.5 py-1.5 rounded-lg bg-white/10 text-white text-xs font-medium border border-white/15 focus:outline-none focus:ring-2 focus:ring-white/30"
          >
            <option value={300} className="text-stone-900">300 dpi</option>
            <option value={203} className="text-stone-900">203 dpi</option>
          </select>
        </div>

        <div className="flex-1 min-w-[100px]" />

        <div className="flex items-end gap-2 flex-wrap">
          {/* 테스트 — 사이즈/여백 점검 (보조) */}
          <button
            type="button"
            onClick={() => handlePrint('test')}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/15 text-white text-xs font-medium hover:bg-white/25 border border-white/10"
            title="첫 번째 라벨 1장만 인쇄 — 사이즈/여백 점검용"
          >
            <Printer size={13} /> 테스트 1장
          </button>
          {/* PRIMARY — 라벨프린터 직접 인쇄 (B방식 1클릭) */}
          <button
            type="button"
            onClick={() => handlePrint('all')}
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg bg-amber-400 text-stone-900 text-xs font-bold hover:bg-amber-300"
            title="라벨프린터로 바로 인쇄 — 다이얼로그에서 [인쇄] 1번. 'PDF로 저장' 선택 시 PDF."
          >
            <Printer size={13} /> 라벨 {items.length}장 인쇄
          </button>
          {/* ZPL — Zebra 네이티브 라벨 파일 저장 (오늘 프린터 없이 prep) */}
          <button
            type="button"
            onClick={handleZpl}
            disabled={zplBusy}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white text-stone-900 text-xs font-bold hover:bg-stone-100 disabled:opacity-50"
            title="Zebra 라벨프린터 네이티브 .zpl 파일로 저장 — 프린터 없이 labelary.com에서 미리보기 가능"
          >
            <FileCode size={13} /> {zplBusy ? '생성 중…' : 'ZPL 저장'}
          </button>
          {/* CSV — 정렬 틀어질 때 보조 경로 */}
          <button
            type="button"
            onClick={handleCsv}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/10 text-white/80 text-xs font-medium hover:bg-white/20 border border-white/10"
            title="보조 경로 — NiceLabel에서 데이터 소스로 연결해 자동 N장 인쇄 (정렬이 틀어질 때)"
          >
            <FileDown size={13} /> CSV
          </button>
          <button
            type="button"
            onClick={handlePng}
            disabled={pngBusy}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white/10 text-white/80 text-xs font-medium hover:bg-white/20 border border-white/10 disabled:opacity-50"
            title="현재 미리보기를 PNG 이미지로 저장"
          >
            <FileImage size={13} /> {pngBusy ? '생성 중…' : 'PNG'}
          </button>
        </div>
      </div>

      {/* 권장 경로 — 라벨프린터 직접 인쇄 (B방식 1클릭) */}
      <div className="label-meta flex items-start gap-2 px-3 py-2.5 rounded-lg bg-amber-50 border-2 border-amber-300 text-[11px] text-stone-800 leading-relaxed">
        <Printer size={14} className="text-amber-700 flex-shrink-0 mt-0.5" />
        <div>
          <strong className="text-stone-900">✨ 권장 — 라벨프린터 직접 인쇄 (버튼 1클릭)</strong>
          <ol className="mt-1 ml-4 space-y-0.5 list-decimal">
            <li>
              <strong>(최초 1회)</strong> 라벨프린터를 Windows <strong>기본 프린터</strong>로 지정 +
              드라이버 용지 <strong>{size.w}×{size.h}mm · 여백 없음</strong> 저장
            </li>
            <li>
              위 <strong>[라벨 {items.length}장 인쇄]</strong> → 인쇄 다이얼로그{' '}
              <strong>[인쇄] 1번</strong> = 라벨프린터 바로 출력
            </li>
          </ol>
          <div className="mt-1.5 text-[10px] text-stone-500">
            🖨️ 권장 기종: <strong>Zebra ZD421T (300dpi)</strong> — 열전사+감열 겸용(확장성), 범용 무지 갭롤, QR 선명
          </div>
          <div className="mt-1 text-[10px] text-stone-600 bg-white/60 rounded px-2 py-1 border border-amber-200">
            📁 <strong>프린터 없는 오늘은</strong> 위 <strong>[ZPL 저장]</strong> → <strong>labelary.com/viewer.html</strong> 에 붙여넣어 라벨 실물 모양 미리 확인. (Zebra 네이티브 .zpl — 도착 후 그대로 출력)
          </div>
        </div>
      </div>

      {/* 라벨프린터 드라이버 1회 설정 가이드 (보조) */}
      <details className="label-meta">
        <summary className="cursor-pointer text-[11px] text-stone-500 hover:text-stone-700 select-none">
          ⋯ 라벨프린터 드라이버 용지 설정 (최초 1회) 펼치기
        </summary>
        <div className="mt-2 flex items-start gap-2 px-3 py-2 rounded-lg bg-stone-50 border border-stone-200 text-[11px] text-stone-700 leading-relaxed">
          <AlertCircle size={13} className="text-stone-500 flex-shrink-0 mt-0.5" />
          <div>
            <strong>드라이버 환경설정(또는 인쇄 다이얼로그 &lsquo;설정 더보기&rsquo;)에서 ↓</strong>
            <ul className="mt-1 ml-3 space-y-0.5 list-disc">
              <li><strong>용지 크기</strong>: &lsquo;사용자 지정&rsquo; → <strong>{size.w} × {size.h} mm</strong></li>
              <li><strong>여백</strong>: 없음 / <strong>배율</strong>: 100%</li>
              <li><strong>머리글·바닥글</strong>: <span className="text-rose-600 font-bold">반드시 끄기</span></li>
              <li><strong>배경 그래픽</strong>: 켜기 (QR 출력)</li>
              <li>한 번 저장하면 이후 <strong>[라벨 N장 인쇄] → [인쇄]</strong> 1클릭으로 바로 출력</li>
              <li className="text-stone-500">※ 정렬이 좌측으로 치우치면(영수증 겸용 계열) → 상단 <strong>[CSV]</strong> 받아 NiceLabel 데이터 소스로 연결</li>
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
