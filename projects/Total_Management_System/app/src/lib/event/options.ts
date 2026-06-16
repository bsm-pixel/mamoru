/**
 * EVENT 접수 옵션 택소노미 (고객 폼 + TMS 공통 SSOT)
 * 종류/옵션은 products.tags 의 {event_type, spec, dry_subtype} 와 매핑.
 */
import type { EventCategoryType } from './types';

export const SLICING_ADDON = 20000; // DRY 슬라이싱 가공 추가비
export const LOW_STOCK_THRESHOLD = 3; // 이하면 '마감임박 N자루' 노출 (그 외 수량 숨김)

export interface EventCategoryDef {
  key: EventCategoryType;
  label: string;
  /** 옵션 종류: 'inch'(블런트·장가위) | 'reduction'(틴닝 감모) | 'dry'(DRY 서브) */
  optionKind: 'inch' | 'reduction' | 'dry';
}

export const EVENT_CATEGORIES: EventCategoryDef[] = [
  { key: 'blunt', label: '블런트', optionKind: 'inch' },
  { key: 'thinning', label: '틴닝', optionKind: 'reduction' },
  { key: 'long', label: '장가위', optionKind: 'inch' },
  { key: 'dry', label: 'DRY', optionKind: 'dry' },
];

/** 손잡이 방향 (오른손/왼손) — 접수폼 1차 분류 */
export const HAND_OPTIONS = [
  { value: 'right', label: '오른손' },
  { value: 'left', label: '왼손' },
];

/** 블런트·장가위 인치 옵션 (필요 시 확장) */
export const INCH_OPTIONS = ['5.0', '5.5', '6.0', '6.5', '7.0'];

/** 틴닝 감모량 구간 */
export const REDUCTION_OPTIONS = [
  { value: 'under20', label: '20% 미만' },
  { value: '21to29', label: '21~29% (메인틴닝 적합)' },
  { value: 'over30', label: '30% 이상' },
];

/** DRY 서브타입 */
export const DRY_OPTIONS = [
  { value: 'slicing', label: '슬라이싱', slicing: true },
  { value: 'stroke', label: '스트록', slicing: false },
  { value: 'curve', label: '커브', slicing: false },
];

export const DRY_SLICING_NOTE =
  '스트록=양날이 살아 있어 모발이 밀리기보다 잘리는 느낌 / 슬라이싱 가공=모발이 밀리는 형태로, 마모루가 직접 가공하여 +20,000원이 발생합니다.';

export const CATEGORY_LABEL: Record<EventCategoryType, string> = {
  blunt: '블런트', thinning: '틴닝', long: '장가위', dry: 'DRY',
};

/** 재고 수량 → 고객 노출 라벨 (적을 때만 수량 노출) */
export function stockLabel(stock: number, threshold = LOW_STOCK_THRESHOLD): {
  text: string; tone: 'soldout' | 'low' | 'ok';
} {
  if (stock <= 0) return { text: '품절', tone: 'soldout' };
  if (stock <= threshold) return { text: `마감임박 · ${stock}자루`, tone: 'low' };
  return { text: '구매가능', tone: 'ok' };
}
