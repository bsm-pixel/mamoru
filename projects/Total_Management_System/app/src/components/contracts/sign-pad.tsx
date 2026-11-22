'use client';

import { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';

/**
 * 최종 서명 — 버튼 → 화면 전체 서명판 → [다시 쓰기] [완료] (2026-11-23, t-261123-sign-ux)
 *
 *  - 손으로 쓰는 것은 이 서명 하나뿐(성함·연락처·주소는 키보드 칸).
 *  - 완료하면 작은 서명 그림으로 보이고, 누르면 다시 쓴다.
 *  - 옛 서명 칸(signature-canvas)의 「손가락을 뗀 뒤 선이 칸 끝까지 이어지던」 문제:
 *    터치·마우스 이벤트를 섞어 쓰던 것을 포인터 이벤트 하나로 바꾸고, 손가락을 떼거나(pointerup)
 *    끊기거나(pointercancel) 칸 밖으로 나가면(pointerleave) 그 획을 끝낸다. 획마다 새 선으로 시작.
 *  - 서명 중에는 화면이 움직이지 않게(touch-action:none + 뒤 화면 스크롤 잠금).
 *  - 화면 배율(devicePixelRatio)에 맞춰 그리고, 중간점 곡선으로 선을 부드럽게.
 *  - 저장 그림은 서명 둘레만 잘라 최대 600×240 으로 줄인다(링크 서명 400,000자 제한 안).
 *  - 고객 서명 페이지(projects/contract/page_sign.html)도 같은 방식으로 그린다 — 고치면 둘 다.
 */

interface Props {
  value: string;
  onChange: (dataUrl: string) => void;
  label?: string;
  /** 서명 전에 꼭 할 것(예: 「위 내용을 확인했습니다」 체크)이 남아 있으면 그 문장 — 버튼이 잠긴다 */
  lockedReason?: string | null;
  testId?: string;
}

export function SignPadField({ value, onChange, label = '서명', lockedReason, testId }: Props) {
  const [open, setOpen] = useState(false);
  return (
    <div data-sign-pad={testId}>
      {value ? (
        <button type="button" onClick={() => setOpen(true)} aria-label={`${label} 다시 쓰기`}
          className="w-full min-h-24 border border-neutral-300 rounded-md bg-white flex items-center justify-center relative hover:border-neutral-800 transition" data-sign-preview>
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img src={value} alt={label} className="max-h-20 max-w-[80%] object-contain" />
          <span className="absolute right-2 bottom-1.5 text-[11px] text-neutral-500 underline underline-offset-2">다시 쓰기</span>
        </button>
      ) : (
        <button type="button" disabled={!!lockedReason} onClick={() => setOpen(true)}
          className="w-full min-h-12 rounded-md bg-neutral-900 text-white text-sm font-semibold disabled:bg-neutral-200 disabled:text-neutral-500 transition" data-sign-open>
          {label}하기
        </button>
      )}
      {!value && lockedReason && <p className="text-xs text-neutral-500 mt-1.5">{lockedReason}</p>}
      {open && <SignPadOverlay label={label} onCancel={() => setOpen(false)} onDone={(url) => { onChange(url); setOpen(false); }} />}
    </div>
  );
}

function SignPadOverlay({ label, onCancel, onDone }: { label: string; onCancel: () => void; onDone: (url: string) => void }) {
  const wrapRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pts = useRef<{ x: number; y: number }[]>([]);
  const drawing = useRef(false);
  const box = useRef({ x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity });
  const [hasInk, setHasInk] = useState(false);

  const setup = useCallback(() => {
    const cv = canvasRef.current, wrap = wrapRef.current;
    if (!cv || !wrap) return;
    const r = wrap.getBoundingClientRect();
    const dpr = Math.min(window.devicePixelRatio || 1, 3);
    cv.width = Math.round(r.width * dpr);
    cv.height = Math.round(r.height * dpr);
    cv.style.width = `${r.width}px`;
    cv.style.height = `${r.height}px`;
    const ctx = cv.getContext('2d');
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.lineWidth = 2.6;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#1A1A1A';
    ctx.fillStyle = '#1A1A1A';
    // 크기가 바뀌면(폰을 눕힘) 그린 것은 지워진다 — 다시 쓰게 한다
    box.current = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
  }, []);

  useEffect(() => {
    setup();
    const prev = document.body.style.overflow;
    document.body.style.overflow = 'hidden'; // 서명 중 뒤 화면이 움직이지 않게
    const onResize = () => { setup(); setHasInk(false); };
    window.addEventListener('resize', onResize);
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onCancel(); };
    window.addEventListener('keydown', onKey);
    return () => { document.body.style.overflow = prev; window.removeEventListener('resize', onResize); window.removeEventListener('keydown', onKey); };
  }, [setup, onCancel]);

  const pos = (e: React.PointerEvent) => {
    const r = canvasRef.current!.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };
  const grow = (p: { x: number; y: number }) => {
    const b = box.current;
    b.x0 = Math.min(b.x0, p.x); b.y0 = Math.min(b.y0, p.y); b.x1 = Math.max(b.x1, p.x); b.y1 = Math.max(b.y1, p.y);
  };

  const down = (e: React.PointerEvent) => {
    e.preventDefault();
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    canvasRef.current!.setPointerCapture(e.pointerId);
    drawing.current = true;
    const p = pos(e);
    pts.current = [p];
    grow(p);
    ctx.beginPath(); ctx.arc(p.x, p.y, 1.3, 0, Math.PI * 2); ctx.fill(); // 점만 찍어도 남게
    setHasInk(true);
  };
  const move = (e: React.PointerEvent) => {
    if (!drawing.current) return;
    e.preventDefault();
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    const p = pos(e);
    const cvr = canvasRef.current!.getBoundingClientRect();
    if (p.x < 0 || p.y < 0 || p.x > cvr.width || p.y > cvr.height) { up(e); return; } // 칸 밖으로 나가면 획 끝
    const a = pts.current;
    a.push(p);
    grow(p);
    if (a.length < 3) {
      ctx.beginPath(); ctx.moveTo(a[0].x, a[0].y); ctx.lineTo(p.x, p.y); ctx.stroke();
      return;
    }
    // 중간점 곡선: 앞 두 점의 가운데 → 지난 점(조절점) → 지금 두 점의 가운데
    const p0 = a[a.length - 3], p1 = a[a.length - 2], p2 = p;
    const m1 = { x: (p0.x + p1.x) / 2, y: (p0.y + p1.y) / 2 };
    const m2 = { x: (p1.x + p2.x) / 2, y: (p1.y + p2.y) / 2 };
    ctx.beginPath(); ctx.moveTo(m1.x, m1.y); ctx.quadraticCurveTo(p1.x, p1.y, m2.x, m2.y); ctx.stroke();
  };
  const up = (e: React.PointerEvent) => {
    if (!drawing.current) return;
    drawing.current = false; // 손가락을 떼면 획 끝 — 다음 획은 새 선으로 시작
    pts.current = [];
    try { canvasRef.current?.releasePointerCapture(e.pointerId); } catch { /* 이미 풀림 */ }
  };

  const clear = () => {
    const cv = canvasRef.current;
    const ctx = cv?.getContext('2d');
    if (!cv || !ctx) return;
    ctx.save(); ctx.setTransform(1, 0, 0, 1, 0, 0); ctx.clearRect(0, 0, cv.width, cv.height); ctx.restore();
    box.current = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
    setHasInk(false);
  };

  const done = () => {
    const cv = canvasRef.current;
    if (!cv || !hasInk) return;
    onDone(exportSignature(cv, box.current));
  };

  return createPortal(
    <div className="fixed inset-0 z-[200] bg-white flex flex-col" role="dialog" aria-modal="true" aria-label={label} data-sign-overlay>
      <div className="flex items-center justify-between px-4 h-12 border-b border-neutral-200 shrink-0">
        <button type="button" onClick={onCancel} className="min-h-10 px-2 text-sm text-neutral-600">닫기</button>
        <p className="text-sm font-semibold text-neutral-900">{label}</p>
        <span className="w-10" />
      </div>
      <div ref={wrapRef} className="relative flex-1 touch-none select-none overscroll-none">
        <canvas ref={canvasRef} className="absolute inset-0 touch-none cursor-crosshair"
          onPointerDown={down} onPointerMove={move} onPointerUp={up} onPointerCancel={up} onPointerLeave={up} />
        {!hasInk && (
          <p className="absolute inset-0 flex flex-col items-center justify-center gap-1 text-neutral-400 pointer-events-none">
            <span className="text-base">여기에 손가락으로 서명해 주세요</span>
            <span className="text-xs">폰을 눕히면 더 넓게 쓸 수 있어요</span>
          </p>
        )}
        <div className="absolute left-8 right-8 bottom-[22%] border-b border-neutral-200 pointer-events-none" />
      </div>
      <div className="grid grid-cols-2 gap-2 p-3 border-t border-neutral-200 shrink-0" style={{ paddingBottom: 'max(12px, env(safe-area-inset-bottom))' }}>
        <button type="button" onClick={clear} disabled={!hasInk} className="min-h-12 rounded-md border border-neutral-300 text-sm font-semibold text-neutral-700 disabled:opacity-40" data-sign-clear>다시 쓰기</button>
        <button type="button" onClick={done} disabled={!hasInk} className="min-h-12 rounded-md bg-neutral-900 text-white text-sm font-semibold disabled:bg-neutral-200 disabled:text-neutral-500" data-sign-done>완료</button>
      </div>
    </div>,
    document.body,
  );
}

/** 서명 둘레만 잘라 최대 600×240 으로 — 큰 화면 서명이 너무 큰 그림이 되지 않게 */
function exportSignature(cv: HTMLCanvasElement, b: { x0: number; y0: number; x1: number; y1: number }): string {
  const dpr = cv.width / (parseFloat(cv.style.width) || cv.width);
  const pad = 12;
  const sx = Math.max(0, (b.x0 - pad) * dpr), sy = Math.max(0, (b.y0 - pad) * dpr);
  const sw = Math.min(cv.width - sx, (b.x1 - b.x0 + pad * 2) * dpr), sh = Math.min(cv.height - sy, (b.y1 - b.y0 + pad * 2) * dpr);
  const scale = Math.min(1, 600 / sw, 240 / sh);
  const out = document.createElement('canvas');
  out.width = Math.max(1, Math.round(sw * scale));
  out.height = Math.max(1, Math.round(sh * scale));
  out.getContext('2d')!.drawImage(cv, sx, sy, sw, sh, 0, 0, out.width, out.height);
  return out.toDataURL('image/png');
}
