'use client';

import { type MarkV2, colorOf, FLAG_COLOR } from './inspection-marks';

/**
 * 읽기 전용 마킹 표시 — 고객이 보는 형태 (수리내역서 미리보기 / page_as_report 참조).
 *  - 점: ✓ 의 꼭지점이 찍은 지점에 정확히
 *  - 선: 드래그 범위
 *  - 플래그: 우측 상단 표기
 */
export function MarkOverlay({ photoUrl, marks, flags }: { photoUrl: string; marks: MarkV2[]; flags: string[] }) {
  if (!photoUrl) {
    return (
      <div className="aspect-[3/4] w-full rounded-xl bg-neutral-100 flex items-center justify-center text-sm text-neutral-400">
        사진 없음
      </div>
    );
  }
  const lines = marks.filter((m) => m.x2 != null && m.y2 != null);
  const points = marks.filter((m) => m.x2 == null);

  return (
    <div className="relative w-full">
      {/* eslint-disable-next-line @next/next/no-img-element */}
      <img src={photoUrl} alt="가위 사진" className="w-full block rounded-xl" draggable={false} />

      <svg viewBox="0 0 100 100" preserveAspectRatio="none" className="absolute inset-0 w-full h-full pointer-events-none">
        {lines.map((m, i) => (
          <g key={i}>
            <line x1={m.x * 100} y1={m.y * 100} x2={m.x2! * 100} y2={m.y2! * 100} stroke="#fff" strokeWidth={6} vectorEffect="non-scaling-stroke" strokeLinecap="round" />
            <line x1={m.x * 100} y1={m.y * 100} x2={m.x2! * 100} y2={m.y2! * 100} stroke={colorOf(m.label)} strokeWidth={3.5} vectorEffect="non-scaling-stroke" strokeLinecap="round" />
          </g>
        ))}
      </svg>

      {points.map((m, i) => (
        <div key={i} className="absolute pointer-events-none" style={{ left: `${m.x * 100}%`, top: `${m.y * 100}%`, width: 0, height: 0 }}>
          <svg width={26} height={26} viewBox="0 0 26 26" style={{ position: 'absolute', left: -11, top: -20, overflow: 'visible', filter: 'drop-shadow(0 1px 2px rgba(0,0,0,0.35))' }}>
            <path d="M5 13 L11 20 L21 6" stroke="#fff" strokeWidth={6} fill="none" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M5 13 L11 20 L21 6" stroke={colorOf(m.label)} strokeWidth={3.5} fill="none" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </div>
      ))}

      {flags.length > 0 && (
        <div className="absolute top-1.5 right-1.5 flex flex-col items-end gap-1">
          {flags.map((n) => (
            <span key={n} className="text-[10px] font-bold text-white px-2 py-0.5 rounded shadow" style={{ background: FLAG_COLOR }}>
              ⚠ {n}
            </span>
          ))}
        </div>
      )}
    </div>
  );
}
