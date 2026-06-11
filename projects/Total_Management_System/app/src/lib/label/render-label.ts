'use client';

/**
 * 라벨 렌더러 — 템플릿+데이터 → 캔버스 합성(브랜드 폰트 + jsbarcode) → 미리보기 PNG + ZPL(^GFA 전체 래스터).
 * 전체 래스터라 미리보기 = 실제 출력(WYSIWYG). 브라우저 전용(canvas/document).
 */

import JsBarcode from 'jsbarcode';
import { mmToDots, buildLabel } from './zpl';
import type { LabelTemplate, LabelData } from './templates';

const FONT_LINK_ID = 'mamoru-label-fonts';

/** 라벨용 브랜드 폰트 로드 보장 (Outfit/Plus Jakarta/Noto Sans KR) */
export async function ensureLabelFonts(): Promise<void> {
  if (typeof document === 'undefined') return;
  if (!document.getElementById(FONT_LINK_ID)) {
    const link = document.createElement('link');
    link.id = FONT_LINK_ID;
    link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Outfit:wght@700;800;900&family=Noto+Sans+KR:wght@400;500;700&display=swap';
    document.head.appendChild(link);
  }
  try {
    await Promise.all([
      (document as Document).fonts.load("700 24px 'Inter'"),
      (document as Document).fonts.load("400 24px 'Inter'"),
      (document as Document).fonts.load("800 24px 'Outfit'"),
      (document as Document).fonts.load("500 24px 'Noto Sans KR'"),
    ]);
    await (document as Document).fonts.ready;
  } catch { /* fallback to default fonts */ }
}

function resolveToken(s: string, data: LabelData): string {
  return String(s ?? '').replace(/\{(\w+)\}/g, (_, k) => (data as Record<string, string | undefined>)[k] ?? '');
}

// 대문자높이(mm) → 폰트 px (cap height ≈ 0.7 * font size)
function fontPx(sizeMm: number, dotsPerMm: number): number {
  return Math.max(6, Math.round((sizeMm / 0.7) * dotsPerMm));
}

// 이미지(로고) 캐시 로드 — 같은 origin(/labels/...)이라 canvas taint 없음
const imgCache = new Map<string, HTMLImageElement>();
function loadImg(src: string): Promise<HTMLImageElement> {
  const cached = imgCache.get(src);
  if (cached) return Promise.resolve(cached);
  return new Promise((res, rej) => {
    const im = new Image();
    im.onload = () => { imgCache.set(src, im); res(im); };
    im.onerror = () => rej(new Error('image load fail: ' + src));
    im.src = src;
  });
}

/** 템플릿+데이터를 캔버스에 합성 (바코드·로고 포함 — 전체 래스터) */
export async function composeLabelCanvas(tpl: LabelTemplate, data: LabelData, dpi: number): Promise<HTMLCanvasElement> {
  const dotsPerMm = dpi / 25.4;
  const W = mmToDots(tpl.widthMm, dpi);
  const H = mmToDots(tpl.heightMm, dpi);
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d');
  if (!ctx) return canvas;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, W, H);
  ctx.imageSmoothingEnabled = false;
  ctx.textBaseline = 'top';

  for (const el of tpl.elements) {
    const x = Math.round(el.xMm * dotsPerMm);
    const y = Math.round(el.yMm * dotsPerMm);

    if (el.kind === 'text') {
      const text = resolveToken(el.text || '', data);
      if (!text) continue;
      const px = fontPx(el.sizeMm || 2, dotsPerMm);
      ctx.fillStyle = '#000';
      ctx.font = `${el.weight || 700} ${px}px '${el.fontFamily || 'Plus Jakarta Sans'}', 'Noto Sans KR', sans-serif`;
      const ctxAny = ctx as CanvasRenderingContext2D & { letterSpacing?: string };
      if (el.letterSpacingPx != null && 'letterSpacing' in ctx) ctxAny.letterSpacing = `${Math.round(el.letterSpacingPx * dotsPerMm / 3.78)}px`;
      ctx.fillText(text, x, y);
      if ('letterSpacing' in ctx) ctxAny.letterSpacing = '0px';
    } else if (el.kind === 'rule') {
      const len = Math.round((el.lengthMm || 10) * dotsPerMm);
      const th = Math.max(1, Math.round((el.thicknessMm || 0.4) * dotsPerMm));
      ctx.fillStyle = '#000';
      ctx.fillRect(x, y, len, th);
    } else if (el.kind === 'image') {
      if (!el.src) continue;
      try {
        const im = await loadImg(el.src);
        const w = Math.round((el.widthMm || 12) * dotsPerMm);
        const h = el.heightMm ? Math.round(el.heightMm * dotsPerMm) : Math.round(w * (im.naturalHeight / im.naturalWidth));
        ctx.imageSmoothingEnabled = true; // 로고는 부드럽게 축소
        ctx.drawImage(im, x, y, w, h);
        ctx.imageSmoothingEnabled = false;
      } catch { /* 로고 파일 없음 — 건너뜀 */ }
    } else if (el.kind === 'barcode') {
      const bcData = resolveToken(el.data || '', data).trim();
      if (!bcData) continue;
      const bw = Math.round((el.widthMm || 25) * dotsPerMm);
      const bh = Math.round((el.heightMm || 6) * dotsPerMm);
      const tmp = document.createElement('canvas');
      try {
        JsBarcode(tmp, bcData, {
          format: 'CODE128',
          displayValue: !!el.showText,
          margin: 0,
          height: bh,
          width: 2,
          fontSize: Math.round(fontPx(1.6, dotsPerMm)),
          textMargin: 1,
        });
        ctx.imageSmoothingEnabled = false;
        const drawH = el.showText ? bh + Math.round(fontPx(2.0, dotsPerMm)) : bh;
        ctx.drawImage(tmp, x, y, bw, drawH);
      } catch { /* 잘못된 데이터 — 건너뜀 */ }
    }
  }
  return canvas;
}

/** 캔버스 → ZPL ^GFA (1비트 패킹, 1=흑) */
function canvasToGfaField(canvas: HTMLCanvasElement): string {
  const W = canvas.width, H = canvas.height;
  const ctx = canvas.getContext('2d');
  if (!ctx) return '';
  const data = ctx.getImageData(0, 0, W, H).data;
  const bytesPerRow = Math.ceil(W / 8);
  const bytes: number[] = [];
  for (let y = 0; y < H; y++) {
    for (let b = 0; b < bytesPerRow; b++) {
      let byte = 0;
      for (let bit = 0; bit < 8; bit++) {
        const xx = b * 8 + bit;
        let dark = 0;
        if (xx < W) {
          const idx = (y * W + xx) * 4;
          const a = data[idx + 3];
          const lum = 0.299 * data[idx] + 0.587 * data[idx + 1] + 0.114 * data[idx + 2];
          dark = a > 128 && lum < 128 ? 1 : 0;
        }
        byte = (byte << 1) | dark;
      }
      bytes.push(byte);
    }
  }
  const total = bytes.length;
  const hex = bytes.map((v) => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `^FO0,0^GFA,${total},${total},${bytesPerRow},${hex}^FS`;
}

/** 라벨 1종 ZPL (전체 래스터). copies = 매수 */
export async function renderLabelZpl(tpl: LabelTemplate, data: LabelData, dpi: number, copies = 1): Promise<string> {
  const canvas = await composeLabelCanvas(tpl, data, dpi);
  return buildLabel({ widthMm: tpl.widthMm, heightMm: tpl.heightMm, dpi }, [canvasToGfaField(canvas)], copies);
}

/** 미리보기 PNG dataURL */
export async function renderLabelPreview(tpl: LabelTemplate, data: LabelData, dpi: number): Promise<string> {
  return (await composeLabelCanvas(tpl, data, dpi)).toDataURL('image/png');
}
