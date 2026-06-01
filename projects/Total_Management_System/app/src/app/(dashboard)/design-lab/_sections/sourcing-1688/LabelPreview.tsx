'use client';

import { useMemo, useRef, useState } from 'react';
import QRCode from 'react-qr-code';
import { Printer, FileImage, AlertCircle, FileDown, FileCode } from 'lucide-react';
import { printSourcingLabels } from '@/lib/sourcing/label-print';
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

export function LabelPreview({ po, qrBaseUrl = DEMO_BASE_URL }: { po: DemoPO; qrBaseUrl?: string }) {
  const items = useMemo(() => po.items.filter((it) => it.product_name), [po.items]);
  const [sizeId, setSizeId] = useState<string>('p40x20');
  const [customW, setCustomW] = useState(40);
  const [customH, setCustomH] = useState(20);
  const [printMode, setPrintMode] = useState<'idle' | 'test' | 'all'>('idle');
  const [pngBusy, setPngBusy] = useState(false);
  const [zplBusy, setZplBusy] = useState(false);
  const [dpi, setDpi] = useState<203 | 300>(300); // ZPL 출력 해상도 (ZD421T=300, 일부 모델=203)
  // 품목별 라벨 매수 (동일 라벨 N장 — 한 종류 제품에 여러 장 부착용). 발주 수량과 무관.
  const [copies, setCopies] = useState<Record<string, number>>({});
  const getCopies = (id: string) => Math.max(1, Math.min(99, copies[id] ?? 1));
  const setCopy = (id: string, n: number) =>
    setCopies((c) => ({ ...c, [id]: Math.max(1, Math.min(99, Math.round(n) || 1)) }));
  // 단가(CNY) × 환율 → 한화(KRW). 환율은 상단에서 입력한 po.exchange_rate.
  const krwOf = (cny: number) => Math.round((cny || 0) * (po.exchange_rate || 0));
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
    const targetItems = mode === 'test' ? items.slice(0, 1) : items;
    if (targetItems.length === 0) return;
    // 공용 인쇄 모듈 사용 (DOM 의존 없이 QR 문자열 렌더) — STEP1 칩과 동일 경로
    printSourcingLabels(
      targetItems.map((it) => ({
        sticker_no: it.sticker_no,
        product_name: it.product_name,
        unit_price: it.unit_price,
        qrValue: `${qrBaseUrl}/${it.id}`,
        copies: mode === 'test' ? 1 : getCopies(it.id),
      })),
      { w: size.w, h: size.h },
      po.exchange_rate
    );
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
      'copies',
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
          `${qrBaseUrl}/${it.id}`,
          it.unit_price,
          getCopies(it.id),
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
  // ⚠️ html2canvas는 Tailwind v4 oklch() 색상에서 예외를 던지므로 사용 안 함.
  // 라벨을 네이티브 canvas에 직접 그린다 (QR=라이브 SVG 이미지, 텍스트=fillText 한글 그대로)
  // → oklch·canvas taint 문제 원천 회피 + 1비트 ^GFA로 패킹.
  const handleZpl = async () => {
    if (!printAreaRef.current || items.length === 0) return;
    setZplBusy(true);
    try {
      const dotsPerMm = dpi / 25.4;
      const pw = Math.round(size.w * dotsPerMm); // print width (dots)
      const ll = Math.round(size.h * dotsPerMm); // label length (dots)
      const bytesPerRow = Math.ceil(pw / 8);
      const ptToDots = (pt: number) => Math.round((pt * dpi) / 72);
      const pad = Math.max(2, Math.round(dotsPerMm)); // 약 1mm
      const qrSize = Math.round(Math.min(pw, ll) * 0.8);

      const cards = Array.from(
        printAreaRef.current.querySelectorAll('.label-card')
      ) as HTMLElement[];

      const loadImage = (src: string) =>
        new Promise<HTMLImageElement>((res, rej) => {
          const im = new Image();
          im.onload = () => res(im);
          im.onerror = () => rej(new Error('QR SVG 이미지 로드 실패'));
          im.src = src;
        });

      const blocks: string[] = [];
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const card = cards[i];

        const canvas = document.createElement('canvas');
        canvas.width = pw;
        canvas.height = ll;
        const ctx = canvas.getContext('2d');
        if (!ctx) continue;
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, pw, ll);
        ctx.imageSmoothingEnabled = false;

        // QR — 라이브 SVG 복제 + xmlns 보강 후 이미지로 그림 (inline SVG라 taint 없음)
        const qrX = pad;
        const qrY = Math.round((ll - qrSize) / 2);
        const svgEl = card?.querySelector('svg');
        if (svgEl) {
          const clone = svgEl.cloneNode(true) as SVGElement;
          clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
          const svgStr = new XMLSerializer().serializeToString(clone);
          const url = URL.createObjectURL(
            new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' })
          );
          try {
            const im = await loadImage(url);
            ctx.imageSmoothingEnabled = false;
            ctx.drawImage(im, qrX, qrY, qrSize, qrSize);
          } finally {
            URL.revokeObjectURL(url);
          }
        }

        // 텍스트 — 전부 #000 (열전사 1비트라 회색은 임계값에서 사라질 수 있음)
        const tx = qrX + qrSize + pad;
        const maxW = pw - tx - pad;
        ctx.fillStyle = '#000';
        ctx.textBaseline = 'top';

        const seq = it.sticker_no.split('-').pop() || '000';

        const numPx = ptToDots(numFontPt);
        const namePx = ptToDots(nameFontPt);
        const pricePx = ptToDots(priceFontPt);
        const gap = Math.max(1, Math.round(namePx * 0.25));
        const priceText = `₩${krwOf(it.unit_price).toLocaleString()}`;

        ctx.font = `${namePx}px 'Noto Sans KR', sans-serif`;
        const nameLines = wrapTextChars(ctx, it.product_name, maxW, 2);

        const totalH = numPx + gap + nameLines.length * (namePx + 2) + gap + pricePx;
        let y = Math.max(pad, Math.round((ll - totalH) / 2));

        ctx.font = `800 ${numPx}px 'Noto Sans KR', sans-serif`;
        ctx.fillText(`#${seq}`, tx, y);
        y += numPx + gap;

        ctx.font = `${namePx}px 'Noto Sans KR', sans-serif`;
        for (const ln of nameLines) {
          ctx.fillText(ln, tx, y);
          y += namePx + 2;
        }
        y += gap;

        ctx.font = `700 ${pricePx}px 'Noto Sans KR', sans-serif`;
        ctx.fillText(priceText, tx, y);

        // 1비트 패킹 (1=흑=출력), MSB=좌측 픽셀
        const data = ctx.getImageData(0, 0, pw, ll).data;
        const bytes: number[] = [];
        for (let yy = 0; yy < ll; yy++) {
          for (let b = 0; b < bytesPerRow; b++) {
            let byte = 0;
            for (let bit = 0; bit < 8; bit++) {
              const x = b * 8 + bit;
              let dark = 0;
              if (x < pw) {
                const idx = (yy * pw + x) * 4;
                const alpha = data[idx + 3];
                const lum =
                  0.299 * data[idx] + 0.587 * data[idx + 1] + 0.114 * data[idx + 2];
                dark = alpha > 128 && lum < 128 ? 1 : 0;
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
          `^XA\n^PW${pw}\n^LL${ll}\n^LH0,0\n^FO0,0^GFA,${total},${total},${bytesPerRow},${hex}^FS\n^PQ${getCopies(it.id)}\n^XZ`
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
      alert(`ZPL 생성 실패: ${e instanceof Error ? e.message : String(e)}`);
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
  // 단가·수량·PO 제거 → 번호+품목명에 세로 공간 몰아주고 크게 (40×20 라벨 = 짧은변 20mm 기준 1.0)
  const fontScale = shortSide / 20;
  // 번호 + 품목명 + 가격(한화) 3줄 — 세로 분배
  const numFontPt = Math.max(10, Math.min(18, 11 * fontScale));
  const nameFontPt = Math.max(9, Math.min(16, 10 * fontScale));
  const priceFontPt = Math.max(8, Math.min(14, 9 * fontScale));

  const totalLabels = items.reduce((s, it) => s + getCopies(it.id), 0);

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
            <Printer size={13} /> 라벨 {totalLabels}장 인쇄
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
              위 <strong>[라벨 {totalLabels}장 인쇄]</strong> → 인쇄 다이얼로그{' '}
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
        실제 출력 사이즈: <strong className="text-stone-900">{size.w} × {size.h} mm</strong> · 품목 {items.length}종 · 총{' '}
        <strong className="text-stone-900">{totalLabels}장</strong>
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
                qrBaseUrl={qrBaseUrl}
                krw={krwOf(it.unit_price)}
                labelW={labelW}
                labelH={labelH}
                qrPx={qrPx}
                textColPx={textColPx}
                numFontPt={numFontPt}
                nameFontPt={nameFontPt}
                priceFontPt={priceFontPt}
              />
              {/* 품목별 라벨 매수 (동일 라벨 N장) — 인쇄/ZPL/PNG 캡처 제외 */}
              <div
                data-html2canvas-ignore
                className="mt-1.5 inline-flex items-center gap-1 rounded-lg bg-stone-100 px-1.5 py-1"
              >
                <span className="text-[10px] text-stone-500 font-medium mr-0.5">매수</span>
                <button
                  type="button"
                  onClick={() => setCopy(it.id, getCopies(it.id) - 1)}
                  className="w-5 h-5 rounded bg-white text-stone-700 text-sm font-bold leading-none hover:bg-stone-200"
                >
                  −
                </button>
                <input
                  type="number"
                  min={1}
                  max={99}
                  value={getCopies(it.id)}
                  onChange={(e) => setCopy(it.id, Number(e.target.value))}
                  className="w-9 text-center text-xs font-bold text-stone-900 bg-white rounded border border-stone-200 py-0.5 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={() => setCopy(it.id, getCopies(it.id) + 1)}
                  className="w-5 h-5 rounded bg-white text-stone-700 text-sm font-bold leading-none hover:bg-stone-200"
                >
                  +
                </button>
              </div>
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
  qrBaseUrl,
  krw,
  labelW,
  labelH,
  qrPx,
  textColPx,
  numFontPt,
  nameFontPt,
  priceFontPt,
}: {
  item: DemoPOItem;
  qrBaseUrl: string;
  krw: number;
  labelW: number;
  labelH: number;
  qrPx: number;
  textColPx: number;
  numFontPt: number;
  nameFontPt: number;
  priceFontPt: number;
}) {
  const qrUrl = `${qrBaseUrl}/${item.id}`;
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
            fontWeight: 700,
            marginTop: 2 * SCREEN_SCALE,
            lineHeight: 1.1,
          }}
        >
          ₩{krw.toLocaleString()}
        </div>
      </div>
    </div>
  );
}

// canvas 텍스트 줄바꿈 (한글은 공백이 없어 글자 단위로 끊고, 마지막 줄 초과 시 … 처리)
function wrapTextChars(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxW: number,
  maxLines: number
): string[] {
  const lines: string[] = [];
  let cur = '';
  for (const ch of text) {
    if (ctx.measureText(cur + ch).width > maxW && cur) {
      lines.push(cur);
      cur = ch;
    } else {
      cur += ch;
    }
  }
  if (cur) lines.push(cur);
  if (lines.length <= maxLines) return lines;
  const kept = lines.slice(0, maxLines);
  let last = kept[maxLines - 1];
  while (last && ctx.measureText(last + '…').width > maxW) last = last.slice(0, -1);
  kept[maxLines - 1] = last + '…';
  return kept;
}

