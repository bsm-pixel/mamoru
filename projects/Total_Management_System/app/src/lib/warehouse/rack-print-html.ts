import type { LocationWithProducts } from '@/hooks/use-warehouse';
import { cellGridPos, cellSpan } from '@/lib/warehouse/location-code';

/**
 * 렉 배치도 인쇄 HTML — A4 한 장에 렉 하나 (2026-07-20)
 *
 * 라벨지가 아니다. 렉을 실물 그대로(위가 높은 단) 그려서 렉 옆면·벽에 붙여 두고
 * "이 자리엔 이 제품" 을 눈으로 대조하는 용도.
 *
 * 한 장 보장 = @page margin 0 + .sheet 를 A4 실치수로 고정하고 overflow:hidden.
 * 그 안에서 칸 높이를 미리 계산해(fitCellH) 애초에 넘치지 않게 맞춘다.
 *
 * React 의존이 없는 순수 함수로 둔 이유: 실제 인쇄물을 헤드리스로 렌더해 검증하기 위해.
 */

export interface PrintLevel {
  levelNo: number;
  cells: LocationWithProducts[];
  rows: number;
  cols: number;
  isShelf: boolean;
  isDrawer: boolean;
}
export interface PrintRack {
  rackNo: number;
  label: string | null;
  columns: number;
  levels: PrintLevel[];   // 화면과 동일하게 내림차순(위→아래)
}
export type Orient = 'portrait' | 'landscape';

/** A4 실치수(mm) — cw/ch = 여백 10mm 를 뺀 내용 영역 */
export const A4 = {
  portrait:  { w: 210, h: 297, cw: 190, ch: 277 },
  landscape: { w: 297, h: 210, cw: 277, ch: 190 },
} as const;

const HEADER_MM = 13;   // 제목 + 부제
const FOOTER_MM = 5;    // 하단 안내
const LEVEL_GAP = 2;    // 단 사이 간격
const LEVEL_HEAD = 4.2; // 단 이름 줄
const DRAWER_PAD = 3;   // 수납함 테두리 안쪽 여백

/** 이 렉을 A4 한 장에 넣으려면 칸 높이를 몇 mm 로 해야 하는가 */
export function fitCellH(rack: PrintRack, orient: Orient): number {
  const avail = A4[orient].ch - HEADER_MM - FOOTER_MM;
  let gridRows = 0;
  let overhead = 4;   // 렉 테두리 padding
  for (const lvl of rack.levels) {
    gridRows += lvl.isDrawer ? lvl.rows : 1;
    overhead += LEVEL_HEAD + LEVEL_GAP + (lvl.isDrawer ? DRAWER_PAD : 0);
  }
  if (gridRows === 0) return 20;
  return Math.max(5, Math.min(28, (avail - overhead) / gridRows));
}

export interface PrintOptions {
  orient: Orient;
  showProduct: boolean;
  /** 출력일 — 테스트에서 고정하기 위해 주입 가능 */
  dateText?: string;
}

/** 인쇄용 HTML — 미리보기(iframe)와 실제 인쇄가 같은 문서를 쓴다 */
export function buildRackPrintHtml(racks: PrintRack[], opt: PrintOptions): string {
  const { orient, showProduct } = opt;
  const size = A4[orient];
  const dateText = opt.dateText ?? today();

  const sheets = racks.map((rack) => {
    const cellH = fitCellH(rack, orient);
    const codePt = Math.max(5.5, Math.min(11, cellH));
    const namePt = Math.max(5, Math.min(9, cellH * 0.62));
    const roomy = cellH >= 9;   // 제품명까지 넣을 여유가 있는가
    // 칸 높이에서 코드줄·재고줄·안쪽여백을 빼고 제품명을 몇 줄까지 넣을 수 있는지 (pt→mm = ×0.3528)
    const lineMm = (pt: number) => pt * 1.2 * 0.3528;
    const maxNameLines = Math.max(1, Math.floor((cellH - lineMm(codePt) - lineMm(namePt) - 1.4) / lineMm(namePt)));

    // 수납함은 칸이 좁고 렉·단이 이미 위에 적혀 있어 뒷자리(A1)만 쓴다 — 화면 배치도와 같은 규칙
    const cell = (loc: LocationWithProducts, span: boolean, short: boolean) => {
      const empty = loc.product_count === 0;
      const shown = short ? (loc.code.split('-').pop() || loc.code) : loc.code;
      // 중간 칸을 삭제해도 밀리지 않게 열·행 명시 + 병합 폭(col_span) 반영 (화면 배치도와 동일 규칙)
      const pos = cellGridPos(loc.bin_no, loc.bin_row);
      const colSpan = cellSpan(loc.col_span);
      const place = span ? 'grid-column:1/-1' : `grid-column:${pos.col} / span ${colSpan};grid-row:${pos.row}`;

      // 한 칸에 여러 품목을 몰아 넣는 경우 — 칸 높이가 허락하는 만큼 이름을 더 적는다
      const names = loc.products.slice(0, maxNameLines)
        .map((p) => `<div class="nm">${esc(p.name)}</div>`).join('');
      const rest = loc.product_count - Math.min(loc.product_count, maxNameLines);
      const foot = `<div class="st">${rest > 0 ? `외 ${rest}종 · ` : ''}재고 ${loc.stock_total}</div>`;
      const body = !showProduct || !roomy
        ? ''
        : empty ? '<div class="nm empty">비어 있음</div>' : names + foot;

      const badge = loc.product_count > 1 ? `<span class="bg">${loc.product_count}종</span>` : '';
      return `<div class="cell${empty ? ' e' : ''}" style="${place}">`
        + `<div class="cd"><span class="ct">${esc(shown)}</span>${badge}</div>${body}</div>`;
    };

    const levels = rack.levels.map((lvl) => {
      // 수납함은 칸에 뒷자리만 적히므로, 전체 코드 형식을 단 이름 옆에 한 번 밝혀 둔다
      const prefix = lvl.cells[0] ? lvl.cells[0].code.replace(/-[^-]*$/, '-') : '';
      const shape = lvl.isShelf ? '선반'
        : lvl.isDrawer ? `수납함 ${lvl.rows}행 ${lvl.cols}열 · 코드 ${esc(prefix)}○`
        : `${lvl.cols}칸`;
      const cols = lvl.isDrawer ? lvl.cols : rack.columns;
      const grid = `<div class="grid" style="grid-template-columns:repeat(${cols},1fr)">`
        + lvl.cells.map((c) => cell(c, lvl.isShelf, lvl.isDrawer)).join('') + '</div>';
      return '<div class="lvl">'
        + `<div class="lh"><b>${lvl.levelNo}단</b> <span>${shape}</span></div>`
        + (lvl.isDrawer ? `<div class="drawer">${grid}</div>` : grid)
        + '</div>';
    }).join('');

    const cellCount = rack.levels.reduce((s, l) => s + l.cells.length, 0);
    const used = rack.levels.reduce((s, l) => s + l.cells.filter((c) => c.product_count > 0).length, 0);

    return `<div class="sheet" style="--ch:${cellH.toFixed(2)}mm;--code:${codePt.toFixed(1)}pt;--name:${namePt.toFixed(1)}pt">
      <div class="hd">
        <h1>${rack.rackNo}번 렉${rack.label ? ` · ${esc(rack.label)}` : ''}</h1>
        <div class="sub">${rack.levels.length}단 · 자리 ${cellCount}개 (사용 ${used} / 빈칸 ${cellCount - used}) · 출력 ${dateText}</div>
      </div>
      <div class="rack">${levels}</div>
      <div class="ft">1단 = 맨 아래 (건물 층수와 동일) · 칸 없는 단은 선반 · 행이 여러 개면 수납함</div>
    </div>`;
  }).join('');

  return `<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>창고 배치도${racks.length === 1 ? ` — ${racks[0].rackNo}번 렉` : ''}</title>
<style>
  @page { size: A4 ${orient}; margin: 0; }
  * { box-sizing: border-box; margin: 0; }
  body { font-family: 'Noto Sans KR','Malgun Gothic','Apple SD Gothic Neo',sans-serif; color:#000; background:#fff; }
  /* 한 장 보장 — 시트를 A4 실치수로 고정하고 넘치면 자른다 (칸 높이는 fitCellH 가 미리 맞춤) */
  .sheet { width:${size.w}mm; height:${size.h}mm; padding:10mm; display:flex; flex-direction:column;
           page-break-after:always; break-after:page; overflow:hidden; }
  .sheet:last-child { page-break-after:auto; break-after:auto; }
  .hd { height:${HEADER_MM}mm; }
  h1 { font-size:15pt; font-weight:800; letter-spacing:-0.3px; }
  .sub { font-size:7.5pt; color:#555; margin-top:1mm; }
  .ft { height:${FOOTER_MM}mm; font-size:6.5pt; color:#888; padding-top:1.5mm; }
  .rack { flex:1; min-height:0; border:0.6mm solid #000; border-radius:1.5mm; padding:1.5mm;
          display:flex; flex-direction:column; gap:${LEVEL_GAP}mm; }
  .lvl { display:flex; flex-direction:column; }
  .lh { height:${LEVEL_HEAD}mm; font-size:7pt; color:#333; line-height:${LEVEL_HEAD}mm; }
  .lh b { font-size:8pt; }
  .lh span { color:#999; font-size:6.5pt; margin-left:1mm; }
  .drawer { border:0.4mm solid #666; border-radius:1mm; padding:${DRAWER_PAD / 2}mm; }
  .grid { display:grid; gap:0.8mm; }
  .cell { height:var(--ch); border:0.25mm solid #333; border-radius:0.8mm; padding:0.6mm 1mm;
          overflow:hidden; display:flex; flex-direction:column; justify-content:center; }
  .cell.e { border-style:dashed; border-color:#bbb; }
  .cd { display:flex; align-items:center; gap:1mm; font-family:'Courier New',monospace; font-weight:700;
        font-size:var(--code); line-height:1.05; }
  .ct { white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .bg { margin-left:auto; flex:none; background:#000; color:#fff; border-radius:0.5mm; padding:0 0.7mm;
        font-family:inherit; font-size:calc(var(--name) * 0.85); font-weight:700; }
  .cell.e .cd { color:#999; font-weight:500; }
  .nm { font-size:var(--name); font-weight:700; line-height:1.15; margin-top:0.4mm;
        white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .nm.empty { font-weight:400; color:#bbb; }
  .st { font-size:var(--name); color:#555; line-height:1.1; }
  @media print { body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }
</style></head><body>${sheets}</body></html>`;
}

function esc(s: string): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function today(): string {
  const d = new Date();
  const p = (n: number) => String(n).padStart(2, '0');
  return `${String(d.getFullYear()).slice(2)}.${p(d.getMonth() + 1)}.${p(d.getDate())}`;
}
