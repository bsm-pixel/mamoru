/**
 * ZPL 라벨 생성 코어 — 순수 함수(DOM 의존 없음). Zebra ZT231 등 ZPL 프린터용.
 *
 * 바코드(^BCN Code128)·QR(^BQN)은 ZPL 네이티브 명령 사용(선명·스캔 안정).
 * 한글 텍스트는 ZPL 폰트에 없으므로 호출측에서 래스터(^GFA)로 만들어 fGfa로 주입(lib/label/render-text).
 * 영문/숫자는 fText(^A0N) 사용 가능.
 *
 * 좌표·크기 단위는 dots. mmToDots로 mm→dots 변환. (203dpi=8dots/mm, 300dpi≈11.8)
 */

export interface LabelSize {
  widthMm: number;
  heightMm: number;
  dpi: number; // 203 | 300
}

export function mmToDots(mm: number, dpi: number): number {
  return Math.round((mm * dpi) / 25.4);
}

/** ZPL ^FD 안전 — 제어문자(^,~) 치환 */
function esc(s: string): string {
  return String(s ?? '').replace(/[\^~]/g, ' ');
}

/** Code128 바코드 (showText=true 면 하단 사람이 읽는 숫자 표시) */
export function fBarcode(xDots: number, yDots: number, data: string, heightDots: number, showText = true, moduleWidth = 2): string {
  return `^FO${xDots},${yDots}^BY${moduleWidth}^BCN,${heightDots},${showText ? 'Y' : 'N'},N,N^FD${esc(data)}^FS`;
}

/** QR 코드 (mag=배율 1~10) */
export function fQR(xDots: number, yDots: number, data: string, mag = 4): string {
  return `^FO${xDots},${yDots}^BQN,2,${mag}^FDLA,${esc(data)}^FS`;
}

/** 영문/숫자 텍스트 (^A0N, 한글 불가). hDots=글자높이, wDots=글자폭 */
export function fText(xDots: number, yDots: number, text: string, hDots: number, wDots?: number): string {
  return `^FO${xDots},${yDots}^A0N,${hDots},${wDots ?? hDots}^FD${esc(text)}^FS`;
}

/** 사전 생성된 ^FO..^GFA..^FS 필드(한글 래스터 등) 그대로 주입 */
export function fGfa(field: string): string {
  return field;
}

/**
 * 라벨 1장 ZPL 조립. fields = 위 f* 들이 만든 문자열 배열. copies = 매수.
 * ^MNY = 갭(간극) 라벨 가정(현 운영 라벨). ^LH0,0 원점.
 */
export function buildLabel(size: LabelSize, fields: string[], copies = 1): string {
  const pw = mmToDots(size.widthMm, size.dpi);
  const ll = mmToDots(size.heightMm, size.dpi);
  return [
    '^XA',
    `^PW${pw}`,
    `^LL${ll}`,
    '^LH0,0',
    '^MNY',
    ...fields.filter(Boolean),
    `^PQ${Math.max(1, Math.round(copies))}`,
    '^XZ',
  ].join('\n');
}

/** 여러 라벨 ZPL 이어붙이기 */
export function buildLabels(labels: string[]): string {
  return labels.join('\n');
}
