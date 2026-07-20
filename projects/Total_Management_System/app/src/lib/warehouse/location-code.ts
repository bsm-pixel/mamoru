/**
 * 창고 로케이션 코드 규칙 — 단일 출처 (112 / 113 / 114, 2026-07-18)
 *
 * 한 단(선반)은 세 가지 형태 중 하나다:
 *   1) 선반 통째      칸막이 없음                     → R01-4
 *   2) 단순 칸막이     1행 N열 (세로 칸막이만)          → R01-2-A  … R01-2-F
 *   3) 수납함(서랍)    M행 N열 격자 (가위 보관함 등)     → R01-2-A1 … R01-2-J6
 *
 * 코드 규칙:
 *  · 열 = 알파벳(A,B,C…), 행 = 숫자(1,2,3…)  ← 엑셀 셀 주소와 같은 감각
 *  · **행이 1개뿐이면 행 번호를 붙이지 않는다** (기존 코드 R01-2-A 를 그대로 유지 = 하위호환)
 *  · 렉 번호는 자릿수 고정(R01) — 문자열 정렬이 숫자 순서와 일치 (R10 이 R2 앞에 오는 사고 방지)
 *  · **단은 아래에서 위로 번호 (1단 = 맨 아래)** — 건물 층수와 같은 감각 (115, 2026-07-18 사장님 지적)
 *    화면은 위가 큰 번호가 되도록 역순으로 그린다.
 *    상/중/하 표기는 쓰지 않는다 — 단 수가 3이 아닐 때 규칙이 깨지고 헷갈리기만 함
 */

/** 열 번호(1-based) → 알파벳. 1→A … 26→Z, 27→AA */
export function binLetter(colNo: number): string {
  let n = Math.max(1, Math.floor(colNo));
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/**
 * 자리 주소 코드.
 * @param col  열 (null = 선반 통째)
 * @param row  행 (null 또는 총 행이 1개면 코드에 안 붙음)
 * @param totalRows 그 단의 총 행 수 (1이면 행 번호 생략)
 */
export function makeLocationCode(
  rackNo: number, levelNo: number,
  col?: number | null, row?: number | null, totalRows = 1,
): string {
  const rack = `R${String(rackNo).padStart(2, '0')}`;
  if (col == null) return `${rack}-${levelNo}`;              // 선반
  const rowPart = totalRows > 1 && row != null ? String(row) : '';
  return `${rack}-${levelNo}-${binLetter(col)}${rowPart}`;
}

/**
 * 사람이 읽는 이름.
 * 단은 항상 'N단' (1단 = 맨 아래). 상/중/하 표기는 쓰지 않는다 —
 * 단 수가 3이 아닐 때 규칙이 깨지고, 1단의 의미도 헷갈리게 만들었다.
 */
export function makeLocationLabel(
  rackNo: number, levelNo: number,
  col?: number | null, row?: number | null,
  totalRows = 1,
): string {
  const rack = `${rackNo}번렉`;
  const level = `${levelNo}단`;
  if (col == null) return `${rack} ${level} 선반`;
  if (totalRows > 1 && row != null) return `${rack} ${level} 수납함 ${binLetter(col)}열 ${row}행`;
  return `${rack} ${level} ${binLetter(col)}칸`;
}

/** 정렬 키 — 렉 → 단 → 행 → 열. 수납함도 위에서 아래, 왼쪽에서 오른쪽으로 정렬된다 */
export function locationSortOrder(rackNo: number, levelNo: number, col?: number | null, row?: number | null): number {
  return rackNo * 1_000_000 + levelNo * 10_000 + (row ?? 0) * 100 + (col ?? 0);
}

/**
 * 격자에서 이 칸이 놓일 자리 (1-based). 2026-07-20
 *
 * 칸을 배열 순서대로 흘려보내면 **중간 칸을 하나 삭제했을 때 뒤의 칸이 전부 밀려서**
 * 배치도가 실제 렉과 어긋난다. 지도가 실물과 다르면 지도가 아니므로,
 * 배치도·수정화면·인쇄 세 곳 모두 이 함수로 열/행을 명시 배치한다.
 */
export function cellGridPos(binNo: number | null, binRow: number | null): { col: number; row: number } {
  return { col: Math.max(1, binNo ?? 1), row: Math.max(1, binRow ?? 1) };
}

export interface GeneratedLocation {
  code: string;
  label: string;
  rack_no: number;
  level_no: number;
  bin_no: number | null;   // 열 (null = 선반)
  bin_row: number | null;  // 행 (null = 선반)
  sort_order: number;
}

/** 한 단의 구조 — cols(열) × rows(행). cols=0 이면 칸 없이 선반 */
export interface LevelSpec {
  cols: number;
  rows?: number;
}

/**
 * 렉 하나를 단·행·열로 펼쳐 로케이션 목록 생성.
 *
 *   [{cols:2}, {cols:6}, {cols:10, rows:6}, {cols:0}]
 *   → 1단 2칸 / 2단 6칸 / 3단 수납함 6행10열(60칸) / 4단 선반
 */
export function generateRackLocations(rackNo: number, levels: LevelSpec[]): GeneratedLocation[] {
  const out: GeneratedLocation[] = [];

  levels.forEach((spec, idx) => {
    const lv = idx + 1;
    const cols = Number.isFinite(spec.cols) && spec.cols > 0 ? Math.floor(spec.cols) : 0;
    const rows = cols === 0 ? 0 : Math.max(1, Math.floor(spec.rows ?? 1));

    if (cols === 0) {
      // 선반 통째
      out.push({
        code: makeLocationCode(rackNo, lv, null),
        label: makeLocationLabel(rackNo, lv, null, null),
        rack_no: rackNo, level_no: lv, bin_no: null, bin_row: null,
        sort_order: locationSortOrder(rackNo, lv, null, null),
      });
      return;
    }

    for (let r = 1; r <= rows; r++) {
      for (let c = 1; c <= cols; c++) {
        out.push({
          code: makeLocationCode(rackNo, lv, c, r, rows),
          label: makeLocationLabel(rackNo, lv, c, r, rows),
          rack_no: rackNo, level_no: lv, bin_no: c, bin_row: r,
          sort_order: locationSortOrder(rackNo, lv, c, r),
        });
      }
    }
  });
  return out;
}
