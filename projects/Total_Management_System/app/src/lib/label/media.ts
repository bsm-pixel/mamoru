/**
 * 라벨지(매체) 설정 묶음 — 출력할 때마다 ZPL에 같이 실어 보내, 롤을 바꿔도 프린터 메뉴를 만질 필요가 없게 한다.
 *
 * 2026-10-25 실측(ZT231): 프린터에 남아 있던 이전 라벨지 설정(인쇄 폭 40mm 등) 때문에 유포 꼬리표가 빈 종이로 나옴
 * → 라벨 종류마다 인쇄 방식·진하기·속도를 고정해 보낸다. (인쇄 폭·길이는 buildLabel 이 ^PW/^LL 로 이미 보냄)
 * 롤 교체 후 간격 측정은 프린터 설정 "헤드 닫을 때 캘리브레이션"(HEAD CLOSE: CALIBRATION)이 자동으로 한다.
 */

export interface MediaProfile {
  id: string;
  name: string;
  /** thermal = 감열(리본 없이 감열지) / transfer = 열전사(리본 사용 — 유포·실버 등) */
  method: 'thermal' | 'transfer';
  /** 진하기 0~30 (~SD, 절대값) */
  darkness: number;
  /** 인쇄 속도 inch/s (^PR) — 레진+유포는 느릴수록 잉크가 잘 붙음 */
  speedIps: number;
}

/**
 * 표준 — 제품·시리얼·직접 만든 라벨·정품 QR 등 꼬리표 외 모든 라벨.
 * 2026-10-25 사장님 결정: 앞으로 모든 출력물은 리본(열전사)으로. 진하기·속도는 첫 출력 후 실물 보고 조정.
 * (출력마다 진하기를 같이 보내므로, 꼬리표의 진하기 25가 다음 라벨에 남지 않는다)
 */
export const MEDIA_RIBBON_STANDARD: MediaProfile = {
  id: 'ribbon_standard',
  name: '표준 라벨 (리본)',
  method: 'transfer',
  darkness: 22,
  speedIps: 4,
};

/** 유포 꼬리표 66mm(2열 엇갈림) + 레진 리본 — 실측 진하기 25·속도 3에서 선명 */
export const MEDIA_YUPO_TAG: MediaProfile = {
  id: 'yupo_tag_66',
  name: '유포 꼬리표 66mm (리본)',
  method: 'transfer',
  darkness: 25,
  speedIps: 3,
};

/** ^XA 앞에 붙는 제어 명령 (진하기는 포맷 밖 ~SD 로 절대값 지정) */
export function mediaPrefix(m?: MediaProfile): string[] {
  if (!m) return [];
  const d = Math.max(0, Math.min(30, Math.round(m.darkness)));
  return [`~SD${String(d).padStart(2, '0')}`];
}

/** ^XA 안에 들어가는 명령 (인쇄 방식·속도) */
export function mediaFormat(m?: MediaProfile): string[] {
  if (!m) return [];
  return [`^MT${m.method === 'transfer' ? 'T' : 'D'}`, `^PR${Math.max(2, Math.min(8, Math.round(m.speedIps)))}`];
}
