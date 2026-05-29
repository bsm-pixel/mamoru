'use client';

import { useRef, useState } from 'react';
import { Undo2 } from 'lucide-react';
import { type MarkV2, MARK_TYPES, FLAG_TYPES, colorOf } from './inspection-marks';
import { MarkOverlay } from './mark-overlay';

/**
 * 수리내역서 핀 마킹 보드 v2 (편집)
 *  - 탭 = 점(✓ 꼭지점이 찍은 지점에 정확히) / 드래그 = 선(무뎌짐 범위)
 *  - 마크 탭 = 삭제 · 플래그 = 우측 상단
 *  - controlled-옵션: marks/flags/onMarks/onFlags 주면 상위 상태 공유, 없으면 내부 상태
 */

const clamp01 = (n: number) => Math.min(1, Math.max(0, n));
const DRAG_THRESHOLD = 0.03;

interface Props {
  photoUrl: string;
  marks?: MarkV2[];
  flags?: string[];
  onMarks?: (marks: MarkV2[]) => void;
  onFlags?: (flags: string[]) => void;
  /** 가위 종류 — 특정 유형 노출 제한(빗살 손상=틴닝) */
  scissorType?: string;
}

export function InspectionMarkBoard({ photoUrl, marks: marksProp, flags: flagsProp, onMarks, onFlags, scissorType }: Props) {
  const [internalMarks, setInternalMarks] = useState<MarkV2[]>([]);
  const [internalFlags, setInternalFlags] = useState<string[]>([]);
  const marks = marksProp ?? internalMarks;
  const flags = flagsProp ?? internalFlags;
  const commitMarks: (m: MarkV2[]) => void = onMarks ?? setInternalMarks;
  const commitFlags: (f: string[]) => void = onFlags ?? setInternalFlags;

  const [activeType, setActiveType] = useState(MARK_TYPES[0].label);
  const wrapRef = useRef<HTMLDivElement>(null);
  const downRef = useRef<{ x: number; y: number } | null>(null);
  const movedRef = useRef(false);
  const [preview, setPreview] = useState<{ x: number; y: number; x2: number; y2: number } | null>(null);

  // 종류 제한 반영 + 활성 유형 보정
  const visibleTypes = MARK_TYPES.filter((t) => !t.only || t.only === scissorType);
  const at = visibleTypes.some((t) => t.label === activeType) ? activeType : (visibleTypes[0]?.label || activeType);
  const lineCapable = !!MARK_TYPES.find((t) => t.label === at)?.line;

  const pos = (e: React.PointerEvent) => {
    const r = wrapRef.current!.getBoundingClientRect();
    return { x: clamp01((e.clientX - r.left) / r.width), y: clamp01((e.clientY - r.top) / r.height) };
  };

  // 핵심: pointerdown(손가락 닿는 즉시)에 처리 — touch-action:none 환경에서 click 미발생/탭 취소 영향 0
  //  · 비-선형(찍힘/부품/스토퍼/빗살): 닿는 즉시 점 ✓
  //  · 무뎌짐(lineCapable): down 기록 → 드래그면 선, 가만히 떼면 점
  const onPointerDown = (e: React.PointerEvent) => {
    const markEl = (e.target as HTMLElement).closest('[data-mark]') as HTMLElement | null;
    if (markEl) { // 마크 탭 = 삭제 (click 아닌 pointerdown 으로 — touch-action:none 환경 대응)
      const idx = parseInt(markEl.getAttribute('data-mark-idx') || '-1', 10);
      if (idx >= 0) removeMark(idx);
      return;
    }
    const p = pos(e);
    if (!lineCapable) {
      commitMarks([...marks, { label: at, x: p.x, y: p.y }]);
      downRef.current = null;
      return;
    }
    downRef.current = p;
    movedRef.current = false;
    try { wrapRef.current?.setPointerCapture(e.pointerId); } catch { /* noop */ }
  };
  const onPointerMove = (e: React.PointerEvent) => {
    if (!downRef.current || !lineCapable) return;
    const p = pos(e);
    if (Math.hypot(p.x - downRef.current.x, p.y - downRef.current.y) > DRAG_THRESHOLD) {
      movedRef.current = true;
      setPreview({ x: downRef.current.x, y: downRef.current.y, x2: p.x, y2: p.y });
    }
  };
  const onPointerUp = (e: React.PointerEvent) => {
    if (downRef.current && lineCapable) {
      const start = downRef.current;
      const p = pos(e);
      if (movedRef.current) commitMarks([...marks, { label: at, x: start.x, y: start.y, x2: p.x, y2: p.y }]);
      else commitMarks([...marks, { label: at, x: start.x, y: start.y }]); // 무뎌짐 탭=점
    }
    downRef.current = null;
    movedRef.current = false;
    setPreview(null);
  };

  const removeMark = (idx: number) => commitMarks(marks.filter((_, i) => i !== idx));
  const undoLast = () => commitMarks(marks.slice(0, -1));
  const toggleFlag = (note: string) => commitFlags(flags.includes(note) ? flags.filter((n) => n !== note) : [...flags, note]);

  if (!photoUrl) {
    return <p className="text-xs text-neutral-400">사진을 먼저 촬영/업로드하면 상처를 표시할 수 있습니다.</p>;
  }

  const counts: Record<string, number> = {};
  marks.forEach((m) => { counts[m.label] = (counts[m.label] || 0) + 1; });

  return (
    <div className="space-y-2">
      {/* 위치형 유형 칩 */}
      <div className="flex gap-1.5 flex-wrap">
        {visibleTypes.map((t) => {
          const active = at === t.label;
          return (
            <button key={t.label} type="button" onClick={() => setActiveType(t.label)}
              className="px-2.5 py-1 text-xs rounded-full border font-medium transition flex items-center gap-1.5"
              style={active ? { background: t.color, borderColor: t.color, color: '#fff' } : { background: '#fff', borderColor: '#e5e5e5', color: '#525252' }}>
              <span className="w-2 h-2 rounded-full" style={{ background: active ? '#fff' : t.color }} />
              {t.label}{t.hint ? <span className="opacity-70 text-[10px]">·{t.hint}</span> : null}
            </button>
          );
        })}
      </div>

      {/* 사진 + 마킹 (편집) */}
      <div ref={wrapRef} onPointerDown={onPointerDown} onPointerMove={onPointerMove} onPointerUp={onPointerUp}
        onPointerCancel={() => { downRef.current = null; movedRef.current = false; setPreview(null); }}
        className="relative w-full select-none rounded-lg overflow-hidden border border-neutral-200 cursor-crosshair"
        style={{ touchAction: 'none' }}>
        <MarkOverlay photoUrl={photoUrl} marks={marks} flags={flags} />

        {/* 드래그 미리보기 */}
        {preview && (
          <svg viewBox="0 0 100 100" preserveAspectRatio="none" className="absolute inset-0 w-full h-full pointer-events-none">
            <line x1={preview.x * 100} y1={preview.y * 100} x2={preview.x2 * 100} y2={preview.y2 * 100} stroke={colorOf(at)} strokeWidth={3} strokeDasharray="4 3" vectorEffect="non-scaling-stroke" strokeLinecap="round" opacity={0.85} />
          </svg>
        )}

        {/* 삭제 히트 영역 */}
        {marks.map((m, i) => {
          if (m.x2 != null && m.y2 != null) {
            const mx = (m.x + m.x2) / 2, my = (m.y + m.y2) / 2;
            return (
              <button key={`l${i}`} data-mark data-mark-idx={i} type="button" title={`${m.label} (탭 삭제)`}
                className="absolute -translate-x-1/2 -translate-y-1/2 w-5 h-5 rounded-full border-2 border-white text-white text-[11px] leading-none flex items-center justify-center shadow"
                style={{ left: `${mx * 100}%`, top: `${my * 100}%`, background: colorOf(m.label) }}>×</button>
            );
          }
          return (
            <button key={`p${i}`} data-mark data-mark-idx={i} type="button" title={`${m.label} (탭 삭제)`}
              className="absolute -translate-x-1/2 -translate-y-1/2 w-8 h-8 rounded-full"
              style={{ left: `${m.x * 100}%`, top: `${m.y * 100}%`, background: 'transparent' }} />
          );
        })}
      </div>

      {/* 플래그 칩 */}
      <div className="flex gap-1.5 flex-wrap">
        {FLAG_TYPES.map((f) => {
          const on = flags.includes(f.note);
          return (
            <button key={f.label} type="button" onClick={() => toggleFlag(f.note)}
              className={`px-2.5 py-1 text-xs rounded-lg border font-medium transition ${on ? 'bg-amber-700 text-white border-amber-700' : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'}`}>
              ⚠ {f.label}
            </button>
          );
        })}
      </div>

      {/* 안내 + 되돌리기 + 범례 */}
      <div className="flex items-center justify-between">
        <span className="text-[11px] text-neutral-400">탭=점 · 드래그=범위(선) · 마크 탭=삭제 ({marks.length})</span>
        <button type="button" onClick={undoLast} disabled={marks.length === 0} className="flex items-center gap-1 text-[11px] text-neutral-500 hover:text-neutral-700 disabled:opacity-40">
          <Undo2 size={12} /> 되돌리기
        </button>
      </div>
      {Object.keys(counts).length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {Object.entries(counts).map(([label, n]) => (
            <span key={label} className="flex items-center gap-1.5 text-[11px] text-neutral-600 bg-neutral-50 border border-neutral-200 rounded-full px-2 py-0.5">
              <span className="w-2.5 h-2.5 rounded-full" style={{ background: colorOf(label) }} />
              {label} {n}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}
