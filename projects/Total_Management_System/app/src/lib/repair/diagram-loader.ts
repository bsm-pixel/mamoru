/**
 * 수리내역서 v3 그림체 로더 — 피그마 SVG 를 불러와 칸(view)별로 잘라 표시할 문자열을 만든다.
 *
 * ⚠️ 고객 페이지(projects/as/page_as_report.html)에 같은 로직의 순수 JS 사본이 있다. 수정 시 두 곳 동기화.
 *
 * - 그림 속 가위 이미지(webp)는 data URL 로 바꿔 박는다
 *   → 상대경로가 TMS 도메인 기준으로 풀리는 문제 + 고객 페이지 이미지 저장(html2canvas)에서 이미지가 빠지는 문제 동시 해결
 * - 칸 범위 = 피그마 프레임(view-blade / view-handle)의 clipPath 사각형
 * - 같은 화면에 그림을 여러 번 그리므로 id(클립·패턴)는 그릴 때마다 고유 접미사를 붙인다
 */

import { DIAGRAM_BASE, type DiagramView } from './wound-catalog';

export interface DiagramSource {
  svg: string;
  views: Partial<Record<DiagramView, string>>;
  layers: string[];
}

const cache = new Map<string, Promise<DiagramSource | null>>();

function blobToDataUrl(b: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result));
    r.onerror = reject;
    r.readAsDataURL(b);
  });
}

function viewBoxOf(doc: Document, id: string): string | null {
  const g = doc.getElementById(id);
  const clip = g?.getAttribute('clip-path')?.match(/url\(#([^)]+)\)/)?.[1];
  const rect = clip ? doc.getElementById(clip)?.querySelector('rect') : null;
  if (!rect) return null;
  const w = rect.getAttribute('width'), h = rect.getAttribute('height');
  const t = rect.getAttribute('transform')?.match(/translate\(\s*([-\d.]+)[ ,]+([-\d.]+)\s*\)/);
  const x = t ? t[1] : rect.getAttribute('x') || '0';
  const y = t ? t[2] : rect.getAttribute('y') || '0';
  return w && h ? `${x} ${y} ${w} ${h}` : null;
}

async function fetchDiagram(file: string): Promise<DiagramSource | null> {
  const res = await fetch(`${DIAGRAM_BASE}${file}.svg`);
  if (!res.ok) return null; // 아직 그림이 없는 종류 → 목록만 표시
  let svg = await res.text();
  const refs = Array.from(new Set(Array.from(svg.matchAll(/href="([^"#][^"]*)"/g)).map((m) => m[1]).filter((u) => !u.startsWith('data:'))));
  for (const ref of refs) {
    const img = await fetch(new URL(ref, `${DIAGRAM_BASE}${file}.svg`).href);
    if (!img.ok) continue;
    const data = await blobToDataUrl(await img.blob());
    svg = svg.split(`href="${ref}"`).join(`href="${data}"`);
  }
  const doc = new DOMParser().parseFromString(svg, 'image/svg+xml');
  const views: DiagramSource['views'] = {};
  for (const v of ['blade', 'handle'] as DiagramView[]) { const vb = viewBoxOf(doc, `view-${v}`); if (vb) views[v] = vb; }
  const layers = Array.from(doc.querySelectorAll('[id^="z-"]')).map((el) => el.id);
  return { svg, views, layers };
}

/** 종류별 그림 1회만 받아 캐시. 실패해도 다음 호출에서 다시 시도 */
export function loadDiagram(file: string): Promise<DiagramSource | null> {
  let p = cache.get(file);
  if (!p) {
    p = fetchDiagram(file).catch(() => null);
    cache.set(file, p);
    p.then((r) => { if (!r) cache.delete(file); });
  }
  return p;
}

/** 한 칸을 그릴 SVG 문자열 — on 에 없는 z- 레이어는 숨김 */
export function renderDiagramView(src: DiagramSource, view: DiagramView, on: string[], uid: string): string | null {
  const vb = src.views[view];
  if (!vb) return null;
  const sfx = `__${uid}`;
  let s = src.svg
    .replace(/<svg[^>]*?>/, (tag) => tag.replace(/\s(width|height)="[^"]*"/g, '').replace(/viewBox="[^"]*"/, `viewBox="${vb}"`))
    .replace(/\sid="([^"]+)"/g, (_, id) => ` id="${id}${sfx}"`)
    .replace(/url\(#([^)]+)\)/g, (_, id) => `url(#${id}${sfx})`)
    .replace(/href="#([^"]+)"/g, (_, id) => `href="#${id}${sfx}"`);
  s = s.replace(/\sid="(z-[^"]+?)__[^"]+"/g, (m, id) => (on.includes(id) ? m : `${m} display="none"`));
  return s;
}
