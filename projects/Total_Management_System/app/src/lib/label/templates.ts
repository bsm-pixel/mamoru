/**
 * 라벨 템플릿 정의 — 사장님 피그마 디자인(2026-06-11) 인코딩. 40×20mm.
 * 변수 토큰: {product}=제품명, {sku}=SKU, {serial}=시리얼(MR…). 좌표·크기는 mm.
 * 위치는 1차 추정 — 미리보기 보고 미세조정.
 */

export interface LabelElement {
  kind: 'text' | 'barcode' | 'rule';
  xMm: number;
  yMm: number;
  // text
  text?: string;            // 리터럴 또는 토큰 포함 (예: "S/N : {serial}")
  fontFamily?: string;      // 'Outfit' | 'Plus Jakarta Sans' | 'Noto Sans KR'
  weight?: number;          // 400~900
  sizeMm?: number;          // 대문자 높이 ≈ mm
  letterSpacingPx?: number;
  // barcode
  data?: string;            // 토큰 (예: "{sku}" / "{serial}")
  widthMm?: number;
  heightMm?: number;
  showText?: boolean;       // 바코드 하단 사람이 읽는 값
  // rule(수평선)
  lengthMm?: number;
  thicknessMm?: number;
}

export interface LabelTemplate {
  id: string;
  name: string;
  widthMm: number;
  heightMm: number;
  elements: LabelElement[];
}

export interface LabelData {
  product?: string;
  sku?: string;
  serial?: string;
  price?: string;
}

/** 품목 바코드 라벨 (시리얼 없음, 바코드=SKU) — 디자인 1 */
export const TEMPLATE_PRODUCT: LabelTemplate = {
  id: 'product_40x20',
  name: '품목 바코드 라벨 40×20',
  widthMm: 40,
  heightMm: 20,
  elements: [
    { kind: 'text', xMm: 2.4, yMm: 1.8, text: 'MAMORU', fontFamily: 'Outfit', weight: 800, sizeMm: 2.4, letterSpacingPx: 1 },
    { kind: 'rule', xMm: 2.4, yMm: 5.0, lengthMm: 19, thicknessMm: 0.4 },
    { kind: 'text', xMm: 2.4, yMm: 6.0, text: '{product}', fontFamily: 'Outfit', weight: 800, sizeMm: 4.2 },
    { kind: 'barcode', xMm: 2.4, yMm: 12.0, data: '{sku}', widthMm: 26, heightMm: 6.5, showText: false },
  ],
};

/** 시리얼 라벨 (바코드=시리얼, 출고 스캔 매칭용) — 디자인 2 */
export const TEMPLATE_SERIAL: LabelTemplate = {
  id: 'serial_40x20',
  name: '시리얼 라벨 40×20',
  widthMm: 40,
  heightMm: 20,
  elements: [
    { kind: 'text', xMm: 2.2, yMm: 1.6, text: 'MAMORU', fontFamily: 'Outfit', weight: 800, sizeMm: 1.9, letterSpacingPx: 0.5 },
    { kind: 'rule', xMm: 2.2, yMm: 4.0, lengthMm: 14, thicknessMm: 0.35 },
    { kind: 'text', xMm: 2.2, yMm: 4.8, text: '{product}', fontFamily: 'Outfit', weight: 800, sizeMm: 2.8 },
    { kind: 'text', xMm: 2.2, yMm: 8.4, text: 'S/N : {serial}', fontFamily: 'Plus Jakarta Sans', weight: 600, sizeMm: 1.5 },
    { kind: 'text', xMm: 2.2, yMm: 13.2, text: 'HAND-CALIBRATED', fontFamily: 'Plus Jakarta Sans', weight: 500, sizeMm: 1.2 },
    { kind: 'text', xMm: 2.2, yMm: 15.4, text: 'PASSED BY BSM', fontFamily: 'Plus Jakarta Sans', weight: 500, sizeMm: 1.2 },
    { kind: 'barcode', xMm: 22.5, yMm: 12.6, data: '{serial}', widthMm: 15, heightMm: 5.5, showText: false },
  ],
};

export const LABEL_TEMPLATES: Record<string, LabelTemplate> = {
  [TEMPLATE_PRODUCT.id]: TEMPLATE_PRODUCT,
  [TEMPLATE_SERIAL.id]: TEMPLATE_SERIAL,
};
