'use client';

import { useRef, useState } from 'react';
import { Undo2 } from 'lucide-react';

/** 사진 위 핀 좌표 — x,y 는 0~1 정규화 비율 (표시 크기 무관) */
export interface PhotoMark {
  x: number;
  y: number;
  label: string;
}

/** 상처/작업 유형 (코드 하드코딩 — 사장님 확정 후 문구 조정) */
export const WOUND_TYPES: { label: string; color: string }[] = [
  { label: '무뎌짐', color: '#D97706' },
  { label: '찍힘', color: '#DC2626' },
  { label: '빗살 손상', color: '#7C3AED' },
  { label: '날 보정', color: '#2563EB' },
  { label: '장력 조정', color: '#0D9488' },
  { label: '부품 교체', color: '#EA580C' },
  { label: '스토퍼 교체', color: '#475569' },
];

export function woundColor(label: string): string {
  return WOUND_TYPES.find((w) => w.label === label)?.color || '#1A1A1A';
}

const clamp01 = (n: number) => Math.min(1, Math.max(0, n));

interface Props {
  photoUrl: string;
  marks: PhotoMark[];
  onChange: (marks: PhotoMark[]) => void;
}

export function InspectionPhotoMarker({ photoUrl, marks, onChange }: Props) {
  const [activeType, setActiveType] = useState(WOUND_TYPES[0].label);
  const wrapRef = useRef<HTMLDivElement>(null);

  const addMark = (e: React.MouseEvent<HTMLDivElement>) => {
    const wrap = wrapRef.current;
    if (!wrap) return;
    const rect = wrap.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return;
    const x = clamp01((e.clientX - rect.left) / rect.width);
    const y = clamp01((e.clientY - rect.top) / rect.height);
    onChange([...marks, { x, y, label: activeType }]);
  };

  const removeMark = (idx: number) => {
    onChange(marks.filter((_, i) => i !== idx));
  };

  const undoLast = () => {
    if (marks.length > 0) onChange(marks.slice(0, -1));
  };

  if (!photoUrl) {
    return (
      <p className="text-xs text-neutral-400">
        사진을 먼저 촬영/업로드하면 상처 위치를 핀으로 표시할 수 있습니다.
      </p>
    );
  }

  return (
    <div className="space-y-2">
      {/* 상처 유형 선택 칩 */}
      <div className="flex gap-1.5 flex-wrap">
        {WOUND_TYPES.map((w) => {
          const active = activeType === w.label;
          return (
            <button
              key={w.label}
              type="button"
              onClick={() => setActiveType(w.label)}
              className="px-2.5 py-1 text-xs rounded-full border font-medium transition flex items-center gap-1.5"
              style={
                active
                  ? { background: w.color, borderColor: w.color, color: '#fff' }
                  : { background: '#fff', borderColor: '#e5e5e5', color: '#525252' }
              }
            >
              <span className="w-2 h-2 rounded-full" style={{ background: active ? '#fff' : w.color }} />
              {w.label}
            </button>
          );
        })}
      </div>

      {/* 사진 + 핀 오버레이 */}
      <div
        ref={wrapRef}
        onClick={addMark}
        className="relative w-full select-none rounded-lg overflow-hidden border border-neutral-200 cursor-crosshair"
        style={{ touchAction: 'manipulation' }}
      >
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img src={photoUrl} alt="검수 사진" className="w-full block pointer-events-none" draggable={false} />
        {marks.map((m, i) => (
          <button
            key={i}
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              removeMark(i);
            }}
            title={`${m.label} (탭하면 삭제)`}
            className="absolute -translate-x-1/2 -translate-y-1/2 flex items-center gap-1 group"
            style={{ left: `${m.x * 100}%`, top: `${m.y * 100}%` }}
          >
            <span
              className="w-5 h-5 rounded-full border-2 border-white shadow-md flex items-center justify-center text-[9px] font-bold text-white"
              style={{ background: woundColor(m.label) }}
            >
              {i + 1}
            </span>
            <span
              className="text-[10px] font-semibold px-1.5 py-0.5 rounded text-white whitespace-nowrap shadow-sm"
              style={{ background: woundColor(m.label) }}
            >
              {m.label}
            </span>
          </button>
        ))}
      </div>

      {/* 하단 컨트롤 */}
      <div className="flex items-center justify-between">
        <span className="text-[11px] text-neutral-400">
          사진을 탭해 「{activeType}」 핀을 찍으세요 · 핀을 탭하면 삭제 ({marks.length}개)
        </span>
        <button
          type="button"
          onClick={undoLast}
          disabled={marks.length === 0}
          className="flex items-center gap-1 text-[11px] text-neutral-500 hover:text-neutral-700 disabled:opacity-40"
        >
          <Undo2 size={12} /> 되돌리기
        </button>
      </div>
    </div>
  );
}
