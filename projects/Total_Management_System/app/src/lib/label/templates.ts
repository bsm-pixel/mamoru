/**
 * 라벨 템플릿 정의 — 사장님 피그마 디자인(2026-06-11) 인코딩. 40×20mm.
 * 변수 토큰: {product}=제품명, {sku}=SKU, {serial}=시리얼(MR…). 좌표·크기는 mm.
 * 위치는 1차 추정 — 미리보기 보고 미세조정.
 */

export interface LabelElement {
  kind: 'text' | 'barcode' | 'rule' | 'image';
  xMm: number;
  yMm: number;
  // image
  src?: string;             // 로고 등 고정 이미지 경로 (예: /labels/mamoru-logo.png)
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
    // MAMORU 로고 = Aggressive 폰트 → 디자인 그대로 이미지 임베드
    // 로고(하단 4.28) → 언더라인 → 모델명 → 바코드(상단 12.0) 균등 ~1mm 간격
    { kind: 'image', xMm: 2.4, yMm: 2.8, src: '/labels/mamoru-logo.png', widthMm: 13 },
    { kind: 'rule', xMm: 2.4, yMm: 5.3, lengthMm: 19, thicknessMm: 0.4 },
    { kind: 'text', xMm: 2.4, yMm: 6.7, text: '{product}', fontFamily: 'Inter', weight: 700, sizeMm: 4.2 },
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
    // MAMORU 로고 = Aggressive 폰트 → 디자인 그대로 이미지 임베드. 여백 2mm, 간격 여유
    { kind: 'image', xMm: 2.0, yMm: 2.4, src: '/labels/mamoru-logo.png', widthMm: 11 },
    { kind: 'rule', xMm: 2.0, yMm: 4.6, lengthMm: 13, thicknessMm: 0.35 },
    { kind: 'text', xMm: 2.0, yMm: 5.6, text: '{product}', fontFamily: 'Inter', weight: 700, sizeMm: 2.6 },
    { kind: 'text', xMm: 2.0, yMm: 9.8, text: 'S/N : {serial}', fontFamily: 'Inter', weight: 400, sizeMm: 1.5 },
    { kind: 'text', xMm: 2.0, yMm: 14.0, text: 'HAND-CALIBRATED', fontFamily: 'Inter', weight: 500, sizeMm: 1.15 },
    { kind: 'text', xMm: 2.0, yMm: 16.2, text: 'PASSED BY BSM', fontFamily: 'Inter', weight: 500, sizeMm: 1.15 },
    { kind: 'barcode', xMm: 21.5, yMm: 7.0, data: '{serial}', widthMm: 16.5, heightMm: 6, showText: false },
  ],
};

export const LABEL_TEMPLATES: Record<string, LabelTemplate> = {
  [TEMPLATE_PRODUCT.id]: TEMPLATE_PRODUCT,
  [TEMPLATE_SERIAL.id]: TEMPLATE_SERIAL,
};
