/**
 * 핀/플래그 선택 → 진단 멘트 자동 연동 (블록 단위 add/remove)
 * 데모(RepairReportDemo)·라이브(InspectionForm) 공유.
 */
import { COMMON_MENT, TININ_INTRO } from '@/lib/repair/comment-presets';
import { FLAG_TYPES, type MarkV2 } from './inspection-marks';

/** 위치형 핀 라벨 → 자동 멘트 */
export const PIN_MENT: Record<string, string> = {
  '무뎌짐': COMMON_MENT.dull,
  '찍힘': COMMON_MENT.nick,
  '빗살 손상': COMMON_MENT.comb,
  '부품 문제': COMMON_MENT.parts,
  '스토퍼 문제': COMMON_MENT.stopper,
};

const flagKey = (note: string) => FLAG_TYPES.find((f) => f.note === note)?.key;

/** 플래그 ON 시 삽입할 멘트(틴닝+장력은 인트로 동반) */
const flagMentOn = (note: string, type: string): string[] => {
  const key = flagKey(note);
  if (key === 'tension') return type === '틴닝' ? [TININ_INTRO, COMMON_MENT.tension] : [COMMON_MENT.tension];
  if (key === 'balance') return [COMMON_MENT.balance];
  if (key === 'edgeangle') return [COMMON_MENT.edgeangle];
  return [];
};
const flagMentOff = (note: string): string[] => {
  const key = flagKey(note);
  if (key === 'tension') return [COMMON_MENT.tension];
  if (key === 'balance') return [COMMON_MENT.balance];
  if (key === 'edgeangle') return [COMMON_MENT.edgeangle];
  return [];
};

export const splitBlocks = (c: string) => c.split('\n\n').map((b) => b.trim()).filter(Boolean);
export const addBlocks = (c: string, blocks: string[]) => {
  const cur = splitBlocks(c);
  for (const b of blocks) { const t = b.trim(); if (t && !cur.includes(t)) cur.push(t); }
  return cur.join('\n\n');
};
export const removeBlocks = (c: string, blocks: string[]) => {
  const rm = blocks.map((b) => b.trim());
  return splitBlocks(c).filter((b) => !rm.includes(b)).join('\n\n');
};

/** 핀 배열 변경 시 라벨별 멘트 자동 add/remove */
export function applyMarkMents(comment: string, prev: MarkV2[], next: MarkV2[]): string {
  let c = comment;
  for (const [label, ment] of Object.entries(PIN_MENT)) {
    const had = prev.some((m) => m.label === label);
    const has = next.some((m) => m.label === label);
    if (has && !had) c = addBlocks(c, [ment]);
    if (!has && had) c = removeBlocks(c, [ment]);
  }
  return c;
}

/** 플래그 배열 변경 시 멘트 자동 add/remove (틴닝+장력 특수) */
export function applyFlagMents(comment: string, prev: string[], next: string[], scissorType: string): string {
  let c = comment;
  for (const n of next.filter((x) => !prev.includes(x))) c = addBlocks(c, flagMentOn(n, scissorType));
  for (const n of prev.filter((x) => !next.includes(x))) c = removeBlocks(c, flagMentOff(n));
  return c;
}
