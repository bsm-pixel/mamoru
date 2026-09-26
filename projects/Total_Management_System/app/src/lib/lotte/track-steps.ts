/**
 * 롯데 ALPS 배송 추적 코드 → 화면 단계 (SSOT, 2026-09-26)
 *
 * 실측(2026-09-26, 송장 12건 조회)으로 확정한 코드표:
 *   02 출력(=우리가 송장 발행) · 09 반품취소 · 10 집하(기사 수거) · 12 운송장등록
 *   20 구간발송 · 21 셔틀/구간도착 · 40 배달전 · 41 배달완료 · 45 인수자등록
 *   실제 흐름: 02 → 10 → 12 → 21 → 20 → 21 → 40 → 41 → 45
 *
 * 🚨 여기 오기 전 화면(delivery-tracker)은 `01 접수 / 02·41 집하 / 42·44 배달중 / 91 배달완료` 였다.
 *    · 01·42·44·91 은 **응답에 없는 코드**라 영원히 안 켜졌고,
 *    · 송장만 출력한 건(02)이 '집하'로 칠해져 **기사님이 가져간 것처럼 보였다**(사장님 지적).
 *    · 41(배달완료)이 '집하'에 묶여 있기도 했다.
 *    백엔드(alps-client.ts)는 이미 올바른 표를 쓰고 있었다 — 화면만 옛 코드로 굳어 있었음.
 *
 * 집하 판정(10 또는 20↑)은 alps-client.isPickedUpCode 가 담당한다(자동 출고완료·알림톡의 기준).
 * 여기서는 '보여주기' 단계만 계산한다. 둘의 기준을 갈라놓지 말 것 — 12(운송장등록)는 양쪽 모두 집하로 보지 않는다.
 */

export interface TrackRecord {
  godsStatCd?: string;
  godsStatNm?: string;   // '출력' · '집하' · '배달완료' …
  scanYmd?: string;      // 'YYYYMMDD'
  scanTme?: string;      // 'HHMMSS' (미입력 시 '000000' 또는 '------')
  brnshpNm?: string;     // 지점명 · 배달완료 시 '고객'
  ptnBrnshpNm?: string;  // 배달완료 시 '집앞' 등
  status?: string;       // '운송장 출력을 하였습니다.' 같은 안내 문장
  [key: string]: unknown;
}

export interface TrackStep {
  key: string;
  label: string;
  codes: string[];
}

/** 화면 4단계 — 라벨은 사장님이 보는 말 그대로 */
export const TRACK_STEPS: TrackStep[] = [
  { key: 'printed', label: '송장 생성', codes: ['02', '12'] },
  { key: 'pickup', label: '집하', codes: ['10'] },
  { key: 'transit', label: '배달중', codes: ['20', '21', '40'] },
  { key: 'delivered', label: '배달완료', codes: ['41', '45'] },
];

export const CANCEL_CODE = '09';

/** 가장 멀리 간 단계(index). 없으면 -1 */
export function resolveStepIndex(tracking: TrackRecord[]): number {
  let max = -1;
  for (const rec of tracking) {
    const code = String(rec.godsStatCd || '');
    for (let i = 0; i < TRACK_STEPS.length; i++) {
      if (TRACK_STEPS[i].codes.includes(code) && i > max) max = i;
    }
  }
  return max;
}

/** 'YYYYMMDD' + 'HHMMSS' → ISO(KST). 시각이 없으면 자정으로 둔다 */
export function scanToIso(rec: TrackRecord): string | null {
  const ymd = String(rec.scanYmd || '').replace(/\D/g, '');
  if (ymd.length !== 8) return null;
  const hms = String(rec.scanTme || '').replace(/\D/g, '').padEnd(6, '0').slice(0, 6);
  const iso = `${ymd.slice(0, 4)}-${ymd.slice(4, 6)}-${ymd.slice(6, 8)}T${hms.slice(0, 2)}:${hms.slice(2, 4)}:${hms.slice(4, 6)}+09:00`;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? null : d.toISOString();
}

/** 각 단계에 도달한 '첫' 스캔 시각 — 스텝바 아래 날짜로 쓴다 */
export function stepTimes(tracking: TrackRecord[]): Record<string, string | null> {
  const out: Record<string, string | null> = {};
  for (const step of TRACK_STEPS) {
    const rec = tracking.find((r) => step.codes.includes(String(r.godsStatCd || '')));
    out[step.key] = rec ? scanToIso(rec) : null;
  }
  return out;
}

/**
 * 마지막 스캔 한 줄 — '집하 · 오류천왕(대) · 9/21 14:03'
 * 🚨 필드명 주의: 응답에 statDt/statTm/orgNm 은 **없다**(옛 화면이 이걸 읽어 빈 줄이 나왔다).
 */
export function formatLastScan(tracking: TrackRecord[]): { label: string; place: string; atText: string; message: string } | null {
  if (tracking.length === 0) return null;
  const rec = tracking[tracking.length - 1];
  const place = String(rec.brnshpNm || '').trim();
  const ymd = String(rec.scanYmd || '').replace(/\D/g, '');
  const hms = String(rec.scanTme || '').replace(/\D/g, '');   // '------' 처럼 시각이 비어 오는 레코드가 있다
  const atText = ymd.length === 8
    ? `${Number(ymd.slice(4, 6))}/${Number(ymd.slice(6, 8))}` + (hms.length >= 4 && hms !== '000000' ? ` ${hms.slice(0, 2)}:${hms.slice(2, 4)}` : '')
    : '';
  return {
    label: String(rec.godsStatNm || LOTTE_STATUS_NAMES[String(rec.godsStatCd || '')] || '').trim(),
    place: place === '고객' ? String(rec.ptnBrnshpNm || '고객').trim() : place,
    atText,
    message: String(rec.status || '').trim(),
  };
}

/** 코드 → 이름 (응답에 godsStatNm 이 비어 올 때의 대비) */
export const LOTTE_STATUS_NAMES: Record<string, string> = {
  '02': '출력',
  '09': '반품취소',
  '10': '집하',
  '12': '운송장등록',
  '20': '구간발송',
  '21': '구간도착',
  '40': '배달전',
  '41': '배달완료',
  '45': '인수자등록',
};
