/**
 * 창고 로케이션 코드 규칙 — 단일 출처 (112, 2026-07-18)
 *
 * 자리 주소 형식:  R{렉2자리}-{단}-{칸알파벳}
 *   R01-2-A  = 1번렉 2단(중단) A칸
 *   R03-1    = 3번렉 1단(상단), 칸 미분할
 *
 * 규칙 3가지 (업계 로케이션 코드 표준):
 *  · 자릿수 고정(R01) — 문자열 정렬이 숫자 순서와 일치 (R10 이 R2 앞에 오는 사고 방지)
 *  · 단은 위→아래 번호 (1=상단) — 사람이 눈으로 보는 순서와 일치
 *  · 칸은 알파벳 — 숫자끼리 헷갈리지 않게 축을 분리
 */

/** 칸 번호(1-based) → 알파벳. 1→A … 26→Z, 27→AA */
export function binLetter(binNo: number): string {
  let n = Math.max(1, Math.floor(binNo));
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** 자리 주소 코드 생성. bin 이 없으면 렉-단 까지만 */
export function makeLocationCode(rackNo: number, levelNo: number, binNo?: number | null): string {
  const rack = `R${String(rackNo).padStart(2, '0')}`;
  return binNo == null ? `${rack}-${levelNo}` : `${rack}-${levelNo}-${binLetter(binNo)}`;
}

/** 사람이 읽는 이름. 3단 기준으로 상/중/하 표기, 그 외는 'N단' */
export function makeLocationLabel(rackNo: number, levelNo: number, binNo?: number | null, totalLevels?: number): string {
  const rack = `${rackNo}번렉`;
  const level =
    totalLevels === 3
      ? (['상단', '중단', '하단'][levelNo - 1] ?? `${levelNo}단`)
      : `${levelNo}단`;
  return binNo == null ? `${rack} ${level}` : `${rack} ${level} ${binLetter(binNo)}칸`;
}

/** 정렬 키 — 렉 → 단 → 칸 순. DB sort_order 에 넣어두면 목록/배치도 정렬이 항상 일치 */
export function locationSortOrder(rackNo: number, levelNo: number, binNo?: number | null): number {
  return rackNo * 10000 + levelNo * 100 + (binNo ?? 0);
}

export interface GeneratedLocation {
  code: string;
  label: string;
  rack_no: number;
  level_no: number;
  bin_no: number | null;
  sort_order: number;
}

/**
 * 렉 하나를 단/칸으로 펼쳐 로케이션 목록 생성.
 *
 * levelBins = 단별 칸 수 배열. 실제 렉은 단마다 칸 수가 다르다.
 *   [2, 6, 6, 0, 0]  →  1단 2칸 / 2단 6칸 / 3단 6칸 / 4·5단은 칸 없이 '선반 통째'
 *   칸 수 0 = 그 단은 나누지 않음 → bin_no NULL 한 칸(전 열을 가로지름)
 */
export function generateRackLocations(rackNo: number, levelBins: number[]): GeneratedLocation[] {
  const out: GeneratedLocation[] = [];
  const totalLevels = levelBins.length;
  levelBins.forEach((binsRaw, idx) => {
    const lv = idx + 1;
    const bins = Number.isFinite(binsRaw) && binsRaw > 0 ? Math.floor(binsRaw) : 0;
    if (bins === 0) {
      out.push({
        code: makeLocationCode(rackNo, lv, null),
        label: makeLocationLabel(rackNo, lv, null, totalLevels),
        rack_no: rackNo, level_no: lv, bin_no: null,
        sort_order: locationSortOrder(rackNo, lv, null),
      });
    } else {
      for (let b = 1; b <= bins; b++) {
        out.push({
          code: makeLocationCode(rackNo, lv, b),
          label: makeLocationLabel(rackNo, lv, b, totalLevels),
          rack_no: rackNo, level_no: lv, bin_no: b,
          sort_order: locationSortOrder(rackNo, lv, b),
        });
      }
    }
  });
  return out;
}
