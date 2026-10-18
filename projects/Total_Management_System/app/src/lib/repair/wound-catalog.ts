/**
 * 수리내역서 v3 — 상처 카탈로그 (SSOT)
 *
 * 가위 종류별 표준 그림(SVG, 피그마 제작) 위에 체크한 상처 레이어만 켜고,
 * 그림 아래에 색 네모 + 문구 목록을 보여준다. (2026-10-15 사장님 확정)
 *
 * - 체크 id = 'z-' + key (구간형은 'z-' + seg + '-' + key). DB repair_inspections.diagram_marks 에 저장
 * - 그림 레이어 id 는 체크 id 와 같다. 단, layer 가 지정된 항목은 다른 항목의 표시를 같이 쓴다
 *   (부품 장착 불량 → 부품 노후 원, 스토퍼 미장착 → 스토퍼 원)
 * - view 가 없는 항목(밸런스·날각)은 그림 표시 없이 목록에만 나온다
 * - 고객 페이지(projects/as/page_as_report.html)는 이 파일을 직접 못 읽으므로
 *   report API 가 woundLines/woundLayers/woundViews 결과를 내려준다 → 문구는 여기만 고치면 된다
 */

export type DiagramView = 'blade' | 'handle';

/** 그림 파일 위치 (GitHub Pages). {file}.svg + {file}-blade.webp / {file}-handle.webp */
export const DIAGRAM_BASE = 'https://page.mamoru.kr/projects/as/img/diagram/';

/** 가위 종류 → 그림 파일. 기타는 블런트 그림 공용 */
export const DIAGRAM_FILE: Record<string, string> = {
  블런트: 'blunt',
  틴닝: 'thinning',
  장가위: 'long',
  슬라이싱: 'slicing',
  기타: 'blunt',
};

export const SEGS = ['tip', 'mid', 'root'] as const;
export type Seg = (typeof SEGS)[number];
export const SEG_KO: Record<Seg, string> = { tip: '날끝', mid: '날 중앙', root: '날 안쪽' };

/** 표시색 — 피그마 그림의 레이어 색과 같아야 한다 (빨강·주황=날 복원 / 파랑=조정 / 보라=교체) */
export const WOUND_COLOR = {
  dull: '#CE4545',
  nick: '#F08A24',
  tension: '#2F80ED',
  part: '#9B3FE0',
  plain: '#8A8580',
  /** 빗살 손상 (틴닝) */
  comb: '#CE4545',
} as const;

export interface WoundKind {
  key: string;
  ko: string;
  view?: DiagramView;
  color: string;
  /** 날끝·날 중앙·날 안쪽 3구간 체크 */
  segs?: boolean;
  /** 구간 체크 id — 기본 z-{seg}-{key}. 빗살은 피그마 이름이 z-comb-{seg}-damage */
  segId?: (s: Seg) => string;
  /** 구간이 이어지면 물결(~)로 합침 (범위 상처). 아니면 가운뎃점(·) 나열 (점 상처) */
  range?: boolean;
  /** 구간형 제목 — 기본 "{구간} {ko}" */
  segTitle?: (where: string) => string;
  /** 그림에서 켤 레이어 (다른 항목 표시 공유) */
  layer?: string;
  /** 특정 가위 종류에서만 노출 */
  only?: string[];
  /** 고객 안내 — 첫 줄=결과, 둘째 줄=사용 팁(연하게) */
  desc: string[];
}

export const WOUND_KINDS: WoundKind[] = [
  { key: 'dull', ko: '무뎌짐', view: 'blade', color: WOUND_COLOR.dull, segs: true, range: true, desc: ['무뎌진 부분이 확인되어 복원했습니다'] },
  { key: 'nick', ko: '찍힘', view: 'blade', color: WOUND_COLOR.nick, segs: true, desc: ['찍힌 상처를 경계선까지 다듬어 복원했습니다'] },
  // 2026-10-18 사장님 확정: 빗살은 '손상' 하나 + 3구간 (틴닝만). 레이어 z-comb-{tip|mid|root}-damage
  {
    key: 'comb-damage', ko: '빗살 손상', view: 'blade', color: WOUND_COLOR.comb, only: ['틴닝'], segs: true, range: true,
    segId: (s) => `z-comb-${s}-damage`,
    segTitle: (where) => (where === '날 전체' ? '빗살 전체 손상' : `${where} 빗살 손상`),
    desc: ['빗살부 손상은 완전 복구가 어려워 최선의 복원으로 마무리했습니다'],
  },
  { key: 'pivot-tension', ko: '장력 헐거움', view: 'handle', color: WOUND_COLOR.tension, desc: ['헐거워진 장력을 맞춰 조정했습니다', '받으실 때의 텐션감을 기억해 두셨다가, 추후 헐거워지면 조여서 사용해 주세요'] },
  { key: 'part-worn', ko: '내부 부품 노후', view: 'handle', color: WOUND_COLOR.part, desc: ['노후된 내부 부품을 새 부품으로 교체했습니다'] },
  { key: 'part-fit', ko: '부품 장착 불량', view: 'handle', layer: 'z-part-worn', color: WOUND_COLOR.part, desc: ['내부 부품이 반대로 장착되어 있어 올바르게 장착했습니다'] },
  { key: 'stopper', ko: '스토퍼 노후', view: 'handle', color: WOUND_COLOR.part, desc: ['닳은 스토퍼를 새 스토퍼로 교체했습니다'] },
  { key: 'stopper-missing', ko: '스토퍼 미장착', view: 'handle', layer: 'z-stopper', color: WOUND_COLOR.part, desc: ['스토퍼가 없는 상태여서 새 스토퍼를 장착했습니다'] },
  // 위치 없음 — v2 플래그(밸런스·날각) 승계. 그림 표시 없이 목록만
  { key: 'balance', ko: '밸런스 불균형', color: WOUND_COLOR.plain, desc: ['가위 밸런스를 교정했습니다'] },
  { key: 'edge-angle', ko: '날각 문제', color: WOUND_COLOR.plain, desc: ['날각을 교정해 커트 안정감을 개선했습니다'] },
];

/** 항목의 체크 id 목록 */
export function segIdOf(kind: WoundKind, s: Seg): string {
  return kind.segId ? kind.segId(s) : `z-${s}-${kind.key}`;
}

export function checkIdsOf(kind: WoundKind): string[] {
  return kind.segs ? SEGS.map((s) => segIdOf(kind, s)) : [`z-${kind.key}`];
}

/** 해당 가위 종류에서 고를 수 있는 항목 */
export function kindsFor(scissorType: string): WoundKind[] {
  return WOUND_KINDS.filter((k) => !k.only || k.only.includes(scissorType));
}

const KIND_BY_ID: Record<string, WoundKind> = {};
for (const k of WOUND_KINDS) for (const id of checkIdsOf(k)) KIND_BY_ID[id] = k;

/** 체크 id 정리 — 모르는 id·중복 제거 */
export function normalizeChecks(raw: unknown): string[] {
  if (!Array.isArray(raw)) return [];
  return Array.from(new Set(raw.filter((x): x is string => typeof x === 'string' && !!KIND_BY_ID[x])));
}

/** 그림에서 켤 레이어 id */
export function woundLayers(checks: string[]): string[] {
  const out = new Set<string>();
  for (const id of checks) {
    const k = KIND_BY_ID[id];
    if (k?.view) out.add(k.layer && !k.segs ? k.layer : id);
  }
  return Array.from(out);
}

/** 켜진 레이어별 표시색 — 그림을 그릴 때 이 색으로 덮어쓴다 (피그마에서 칠한 색과 무관, 색 변경 = 여기만) */
export function woundLayerColors(checks: string[]): Record<string, string> {
  const out: Record<string, string> = {};
  for (const id of checks) {
    const k = KIND_BY_ID[id];
    if (k?.view) out[k.layer && !k.segs ? k.layer : id] = k.color;
  }
  return out;
}

/** 보여줄 그림 칸 (체크된 표시가 있는 칸만) */
export function woundViews(checks: string[]): DiagramView[] {
  const v = new Set<DiagramView>();
  for (const id of checks) { const k = KIND_BY_ID[id]; if (k?.view) v.add(k.view); }
  return (['blade', 'handle'] as DiagramView[]).filter((x) => v.has(x));
}

/** 구간 문구: 범위 상처(무뎌짐·빗살) 연속=물결(~), 떨어짐=가운뎃점(·), 셋 다=날 전체. 찍힘(점 상처)은 나열만 */
function segWhere(kind: WoundKind, checks: string[]): string | null {
  const idx = SEGS.map((s, i) => (checks.includes(segIdOf(kind, s)) ? i : -1)).filter((i) => i >= 0);
  if (!idx.length) return null;
  if (idx.length === 3) return '날 전체';
  const names = idx.map((i) => SEG_KO[SEGS[i]]);
  if (kind.range && idx.length === 2 && idx[1] - idx[0] === 1) return `${names[0]}~${names[1]}`;
  return names.join('·');
}

export interface WoundLine {
  key: string;
  title: string;
  desc: string[];
  color: string;
}

/** 고객 목록 — 카탈로그 순서대로, 구간형은 한 줄로 합침 */
export function woundLines(checks: string[]): WoundLine[] {
  const lines: WoundLine[] = [];
  for (const k of WOUND_KINDS) {
    let title: string | null = null;
    if (k.segs) { const w = segWhere(k, checks); title = w ? (k.segTitle ? k.segTitle(w) : `${w} ${k.ko}`) : null; }
    else if (checks.includes(`z-${k.key}`)) title = k.ko;
    if (title) lines.push({ key: k.key, title, desc: k.desc, color: k.color });
  }
  return lines;
}

/** 한 줄 요약 (TMS 검수 결과·준비표) — 예: "날 전체 무뎌짐 · 장력 헐거움" */
export function woundSummary(checks: string[]): string {
  const lines = woundLines(checks);
  return lines.length ? lines.map((l) => l.title).join(' · ') : '상처 없음 (양호)';
}

/** 준비표·목록용 한 줄 — v3 상처 요약 + 추가 안내(comment). 구 내역서는 comment 만 */
export function inspectionNote(i: { diagram_marks?: unknown; comment?: string | null }): string {
  const parts: string[] = [];
  if (i.diagram_marks != null) parts.push(woundSummary(normalizeChecks(i.diagram_marks)));
  if (i.comment) parts.push(i.comment);
  return parts.join(' / ');
}
