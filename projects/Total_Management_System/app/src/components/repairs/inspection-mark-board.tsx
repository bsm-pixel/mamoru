'use client';

import { useRef, useState } from 'react';
import { Undo2 } from 'lucide-react';

/**
 * 수리내역서 핀 마킹 보드 v2 (디자인모니터 데모 전용 — 승인 후 라이브 승격)
 *  - 탭 = 점(문제 지점), 체크(✓)의 꼭지점이 정확히 그 지점에 위치
 *  - 드래그 = 선(무뎌짐처럼 범위로 잡히는 상처: 시작→끝)
 *  - 플래그 = 위치 없이 우측 상단에 표기 (밸런스/날각 등 부위 특정 불가 항목)
 */

interface MarkV2 {
  label: string;
  x: number;
  y: number;
  x2?: number;
  y2?: number;
}

/** 위치형 문제 유형 (사진 위 핀) */
const MARK_TYPES: { label: string; color: string; hint?: string }[] = [
  { label: '무뎌짐', color: '#D97706', hint: '드래그로 범위' },
  { label: '찍힘', color: '#DC2626' },
  { label: '빗살 손상', color: '#7C3AED' },
  { label: '장력조절 필요', color: '#0D9488' },
  { label: '부품교체 필요', color: '#EA580C' },
  { label: '스토퍼 불량', color: '#475569' },
];

/** 우측 상단 플래그 (위치 없음) */
const FLAG_TYPES: { label: string; note: string; color: string }[] = [
  { label: '밸런스 불균형', note: '가위 밸런스 불균형 — 교정 필요', color: '#B45309' },
  { label: '날각 문제', note: '날각 문제 — 날등 각도 개선 필요', color: '#B45309' },
];

function colorOf(label: string): string {
  return MARK_TYPES.find((t) => t.label === label)?.color || '#1A1A1A';
}
const clamp01 = (n: number) => Math.min(1, Math.max(0, n));
const DRAG_THRESHOLD = 0.03; // 이보다 크게 끌면 선(line)

interface Props {
  photoUrl: string;
}

export function InspectionMarkBoard({ photoUrl }: Props) {
  const [marks, setMarks] = useState<MarkV2[]>([]);
  const [flags, setFlags] = useState<string[]>([]);
  const [activeType, setActiveType] = useState(MARK_TYPES[0].label);
  const wrapRef = useRef<HTMLDivElement>(null);
  const dragRef = useRef<{ x: number; y: number } | null>(null);
  const [preview, setPreview] = useState<{ x: number; y: number; x2: number; y2: number } | null>(null);
  // 사진 변경 시 마킹 초기화는 부모에서 key={photoUrl} 로 리마운트 처리

  const pos = (e: React.PointerEvent) => {
    const r = wrapRef.current!.getBoundingClientRect();
    return { x: clamp01((e.clientX - r.left) / r.width), y: clamp01((e.clientY - r.top) / r.height) };
  };

  const onPointerDown = (e: React.PointerEvent) => {
    if ((e.target as HTMLElement).closest('[data-mark]')) return; // 기존 마크 클릭(삭제)은 무시
    wrapRef.current?.setPointerCapture(e.pointerId);
    dragRef.current = pos(e);
  };
  const onPointerMove = (e: React.PointerEvent) => {
    if (!dragRef.current) return;
    const p = pos(e);
    setPreview({ x: dragRef.current.x, y: dragRef.current.y, x2: p.x, y2: p.y });
  };
  const onPointerUp = (e: React.PointerEvent) => {
    if (!dragRef.current) return;
    const start = dragRef.current;
    const p = pos(e);
    dragRef.current = null;
    setPreview(null);
    const dist = Math.hypot(p.x - start.x, p.y - start.y);
    if (dist > DRAG_THRESHOLD) {
      setMarks((m) => [...m, { label: activeType, x: start.x, y: start.y, x2: p.x, y2: p.y }]);
    } else {
      setMarks((m) => [...m, { label: activeType, x: start.x, y: start.y }]);
    }
  };

  const removeMark = (idx: number) => setMarks((m) => m.filter((_, i) => i !== idx));
  const undoLast = () => setMarks((m) => m.slice(0, -1));
  const toggleFlag = (note: string) =>
    setFlags((f) => (f.includes(note) ? f.filter((n) => n !== note) : [...f, note]));

  if (!photoUrl) {
    return <p className="text-xs text-neutral-400">사진을 먼저 촬영/업로드하면 상처를 표시할 수 있습니다.</p>;
  }

  const counts: Record<string, number> = {};
  marks.forEach((m) => { counts[m.label] = (counts[m.label] || 0) + 1; });

  return (
    <div className="space-y-2">
      {/* 위치형 유형 칩 */}
      <div className="flex gap-1.5 flex-wrap">
        {MARK_TYPES.map((t) => {
          const active = activeType === t.label;
          return (
            <button
              key={t.label}
              type="button"
              onClick={() => setActiveType(t.label)}
              className="px-2.5 py-1 text-xs rounded-full border font-medium transition flex items-center gap-1.5"
              style={active ? { background: t.color, borderColor: t.color, color: '#fff' } : { background: '#fff', borderColor: '#e5e5e5', color: '#525252' }}
            >
              <span className="w-2 h-2 rounded-full" style={{ background: active ? '#fff' : t.color }} />
              {t.label}{t.hint ? <span className="opacity-70 text-[10px]">·{t.hint}</span> : null}
            </button>
          );
        })}
      </div>

      {/* 사진 + 마킹 */}
      <div
        ref={wrapRef}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={() => { dragRef.current = null; setPreview(null); }}
        className="relative w-full select-none rounded-lg overflow-hidden border border-neutral-200 cursor-crosshair"
        style={{ touchAction: 'none' }}
      >
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src={photoUrl} alt="검수 사진" className="w-full block pointer-events-none" draggable={false} />

        {/* 선(line) + 미리보기 SVG */}
        <svg viewBox="0 0 100 100" preserveAspectRatio="none" className="absolute inset-0 w-full h-full pointer-events-none">
          {marks.map((m, i) => m.x2 != null && m.y2 != null && (
            <g key={`l${i}`}>
              <line x1={m.x * 100} y1={m.y * 100} x2={m.x2 * 100} y2={m.y2 * 100} stroke="#fff" strokeWidth={6} vectorEffect="non-scaling-stroke" strokeLinecap="round" />
              <line x1={m.x * 100} y1={m.y * 100} x2={m.x2 * 100} y2={m.y2 * 100} stroke={colorOf(m.label)} strokeWidth={3.5} vectorEffect="non-scaling-stroke" strokeLinecap="round" />
            </g>
          ))}
          {preview && (
            <line x1={preview.x * 100} y1={preview.y * 100} x2={preview.x2 * 100} y2={preview.y2 * 100} stroke={colorOf(activeType)} strokeWidth={3} strokeDasharray="4 3" vectorEffect="non-scaling-stroke" strokeLinecap="round" opacity={0.8} />
          )}
        </svg>

        {/* 마크 핸들 (점=✓꼭지점 정확 / 선=중점 × 삭제 핸들) */}
        {marks.map((m, i) => {
          const c = colorOf(m.label);
          if (m.x2 != null && m.y2 != null) {
            const mx = (m.x + m.x2) / 2, my = (m.y + m.y2) / 2;
            return (
              <button
                key={`h${i}`}
                data-mark
                type="button"
                onClick={() => removeMark(i)}
                title={`${m.label} (탭 삭제)`}
                className="absolute -translate-x-1/2 -translate-y-1/2 w-5 h-5 rounded-full border-2 border-white text-white text-[11px] leading-none flex items-center justify-center shadow"
                style={{ left: `${mx * 100}%`, top: `${my * 100}%`, background: c }}
              >×</button>
            );
          }
          // 점: ✓ 의 꼭지점(11,20)이 정확히 (x,y) 에 오도록
          return (
            <button
              key={`p${i}`}
              data-mark
              type="button"
              onClick={() => removeMark(i)}
              title={`${m.label} (탭 삭제)`}
              className="absolute"
              style={{ left: `${m.x * 100}%`, top: `${m.y * 100}%`, width: 0, height: 0 }}
            >
              <svg width={26} height={26} viewBox="0 0 26 26" style={{ position: 'absolute', left: -11, top: -20, overflow: 'visible', filter: 'drop-shadow(0 1px 2px rgba(0,0,0,0.35))' }}>
                <path d="M5 13 L11 20 L21 6" stroke="#fff" strokeWidth={6} fill="none" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M5 13 L11 20 L21 6" stroke={c} strokeWidth={3.5} fill="none" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </button>
          );
        })}

        {/* 우측 상단 플래그 */}
        {flags.length > 0 && (
          <div className="absolute top-1.5 right-1.5 flex flex-col items-end gap-1">
            {flags.map((n) => (
              <span key={n} className="text-[10px] font-bold text-white px-2 py-0.5 rounded shadow" style={{ background: '#B45309' }}>
                ⚠ {n}
              </span>
            ))}
          </div>
        )}
      </div>

      {/* 플래그 칩 (누르면 우측 상단 표기) */}
      <div className="flex gap-1.5 flex-wrap">
        {FLAG_TYPES.map((f) => {
          const on = flags.includes(f.note);
          return (
            <button
              key={f.label}
              type="button"
              onClick={() => toggleFlag(f.note)}
              className={`px-2.5 py-1 text-xs rounded-lg border font-medium transition ${on ? 'bg-amber-700 text-white border-amber-700' : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'}`}
            >
              ⚠ {f.label}
            </button>
          );
        })}
      </div>

      {/* 하단: 안내 + 범례 + 되돌리기 */}
      <div className="flex items-center justify-between">
        <span className="text-[11px] text-neutral-400">
          탭=점 · 드래그=범위(선) · 마크 탭=삭제 ({marks.length})
        </span>
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
