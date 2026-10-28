'use client';

/**
 * 복원수리 꼬리표 라벨 — 가위마다 1장, 입고 때 출력해 손잡이에 감는다(가위 뒤바뀜 방지).
 *
 * 라벨지: 유포 꼬리표 66mm 폭, 한 칸(20.75mm)에 꼬리표 2개 — 왼쪽은 바로, 오른쪽은 180° 뒤집힌 배치.
 * 패드 1장 = 위/아래 반(가운데 접는선). 위 반 = MAMORU 로고 + "김보금 님", 아래 반 = "01 / 03" + 접수번호(짧게).
 *
 * 좌표는 2026-10-25 ZT231(203dpi) 실물 눈금 출력·스캔으로 실측:
 *   왼쪽 패드 중심 x≈11.5mm(92dot) · 오른쪽 패드 중심 x≈56.5mm(452dot) (10-25 실사용 출력 사진으로 재보정) · 접는선(세로 중심) y≈11.75mm(94dot)
 *   패드 폭 ≈15.4mm → 글자 영역은 13mm(104dot)×17.5mm(140dot)로 여유를 둔다.
 * 위치를 고칠 땐 이 숫자만 바꾼다(실물 출력 → 스캔 확인).
 */

import { buildLabel } from './zpl';
import { MEDIA_YUPO_TAG } from './media';
import { canvasToGfaField, ensureLabelFonts } from './render-label';

const DPI = 203;
const W = 528; // 66mm
const H = 166; // 20.75mm (프린터 실측 라벨 길이)
const CY = 94; // 접는선
const CX = [92, 452] as const; // 왼쪽·오른쪽 패드 중심 — 10-25 실사용 사진 실측으로 왼쪽 +1.9mm·오른쪽 +1mm 보정
const BOX_W = 98; // 패드(≈15.4mm)보다 넉넉히 좁게 — 좌우 여백 확보
const BOX_H = 140;
const LOGO_SRC = '/labels/mamoru-logo.png'; // AGRESSIVE 워드마크(라벨용) — 이름 위 폭 9mm
const LOGO_W = 72;
const FONT = "'Pretendard Variable', Pretendard, 'Noto Sans KR', sans-serif";

export interface RepairTagItem {
  name: string; // 고객 실명
  no: number; // 몇 번째 가위
  total: number; // 전체 가위 수
  code: string; // 접수번호(짧게) AS-1025-003
}

/** AS-20261025-003 → AS-1025-003 · AS-20261028-4827 → AS-1028-4827 (꼬리표용 짧은 표기, 한 해 안에서 유일) */
export function shortAsId(asId: string): string {
  const m = /^AS-\d{4}(\d{4})-(\d+)$/.exec(asId || '');
  return m ? `AS-${m[1]}-${m[2]}` : asId || '';
}

/** 접수 1건 → 가위 수만큼 꼬리표 항목 */
export function buildRepairTagItems(name: string, asId: string, count: number): RepairTagItem[] {
  const total = Math.max(1, Math.min(99, Math.round(count)));
  const code = shortAsId(asId);
  return Array.from({ length: total }, (_, i) => ({ name: (name || '').trim(), no: i + 1, total, code }));
}

const pad2 = (n: number) => String(n).padStart(2, '0');

let logoCache: HTMLImageElement | null = null;
async function loadLogo(): Promise<HTMLImageElement | null> {
  if (logoCache) return logoCache;
  return new Promise((res) => {
    const im = new Image();
    im.onload = () => { logoCache = im; res(im); };
    im.onerror = () => res(null); // 로고 없으면 글자만 출력
    im.src = LOGO_SRC;
  });
}

/** 꼬리표 1개를 (cx, CY) 중심에 그린다. rotate=true 면 180° (오른쪽 칸) */
function drawTag(ctx: CanvasRenderingContext2D, item: RepairTagItem, cx: number, rotate: boolean, logo: HTMLImageElement | null) {
  ctx.save();
  ctx.translate(cx, CY);
  if (rotate) ctx.rotate(Math.PI);
  ctx.fillStyle = '#000';
  ctx.textBaseline = 'alphabetic';
  const top = -BOX_H / 2;

  // 맨 위: MAMORU 로고(브랜드 서체 이미지) — 2026-10-25 실물 확인 "선명해 좋아"
  if (logo) {
    const lh = Math.round((logo.naturalHeight * LOGO_W) / logo.naturalWidth);
    ctx.imageSmoothingEnabled = true;
    ctx.drawImage(logo, -LOGO_W / 2, top + 10, LOGO_W, lh);
  }

  // 위 반: 이름(굵게) + " 님"(한 단계 작게). 넘치면 이름 글자를 줄인다(4글자 이상 이름 대비)
  const honor = ' 님';
  ctx.font = `600 19px ${FONT}`;
  const honorW = ctx.measureText(honor).width;
  let size = 26; // 로고 아래 이름 — 넘치면 아래에서 자동 축소 // 이름이 가장 크게(사장님 기준) — 넘치면 아래에서 자동 축소
  ctx.font = `700 ${size}px ${FONT}`;
  let nameW = ctx.measureText(item.name).width;
  while (nameW + honorW > BOX_W - 4 && size > 14) {
    size -= 1;
    ctx.font = `700 ${size}px ${FONT}`;
    nameW = ctx.measureText(item.name).width;
  }
  const x0 = -(nameW + honorW) / 2;
  ctx.textAlign = 'left';
  ctx.fillText(item.name, x0, top + 52);
  ctx.font = `600 19px ${FONT}`;
  ctx.fillText(honor, x0 + nameW, top + 52);

  // 아래 반: 01 / 03 (크게) + 접수번호(작게) — 접는선에서 살짝 띄움
  ctx.textAlign = 'center';
  ctx.font = `700 26px ${FONT}`;
  ctx.fillText(`${pad2(item.no)} / ${pad2(item.total)}`, 0, top + 109); // 사장님 요청 1mm(8dot) 아래로
  // 접수번호 — 2026-10-28 뒤 4자리 임의(AS-1028-8888 ≈ 95px)로 길어져, 칸을 넘으면 한 단계씩 줄인다
  let codeSize = 14;
  ctx.font = `400 ${codeSize}px ${FONT}`;
  while (ctx.measureText(item.code).width > BOX_W - 4 && codeSize > 11) {
    codeSize -= 1;
    ctx.font = `400 ${codeSize}px ${FONT}`;
  }
  ctx.fillText(item.code, 0, top + 134);
  ctx.restore();
}

/** 한 칸(꼬리표 2개) 캔버스 */
function composeUnit(logo: HTMLImageElement | null, left?: RepairTagItem, right?: RepairTagItem): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = W;
  canvas.height = H;
  const ctx = canvas.getContext('2d');
  if (!ctx) return canvas;
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, W, H);
  if (left) drawTag(ctx, left, CX[0], false, logo);
  if (right) drawTag(ctx, right, CX[1], true, logo);
  return canvas;
}

function pairs(items: RepairTagItem[]): [RepairTagItem, RepairTagItem | undefined][] {
  const out: [RepairTagItem, RepairTagItem | undefined][] = [];
  for (let i = 0; i < items.length; i += 2) out.push([items[i], items[i + 1]]);
  return out;
}

async function ensureTagFonts() {
  await ensureLabelFonts(); // Noto Sans KR 폴백 보장
  try {
    await Promise.all([document.fonts.load(`700 26px ${FONT}`), document.fonts.load(`600 19px ${FONT}`)]);
  } catch { /* 폴백 글꼴로 진행 */ }
}

/** 꼬리표 전체 ZPL — 2개씩 한 칸, 칸마다 1포맷(내용이 다르므로) */
export async function renderRepairTagsZpl(items: RepairTagItem[]): Promise<string> {
  await ensureTagFonts();
  const logo = await loadLogo();
  return pairs(items)
    .map(([l, r]) => buildLabel({ widthMm: 66, heightMm: 20.75, dpi: DPI }, [canvasToGfaField(composeUnit(logo, l, r))], 1, MEDIA_YUPO_TAG))
    .join('\n');
}

/** 미리보기 — 칸별 PNG dataURL */
export async function renderRepairTagsPreview(items: RepairTagItem[]): Promise<string[]> {
  await ensureTagFonts();
  const logo = await loadLogo();
  return pairs(items).map(([l, r]) => composeUnit(logo, l, r).toDataURL('image/png'));
}
