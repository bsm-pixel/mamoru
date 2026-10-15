'use client';

import { useEffect, useId, useState } from 'react';
import { DIAGRAM_FILE, woundLayers, woundViews, type DiagramView } from '@/lib/repair/wound-catalog';
import { loadDiagram, renderDiagramView, type DiagramSource } from '@/lib/repair/diagram-loader';

interface Props {
  scissorType: string;
  checks: string[];
  /** 지정 시 그 칸만, 없으면 체크된 표시가 있는 칸만 (고객 화면과 동일) */
  views?: DiagramView[];
  className?: string;
}

/** 수리내역서 v3 그림체 — 체크한 상처 레이어만 켜서 표시 (그림 소스=피그마 SVG, GitHub Pages) */
export function WoundDiagram({ scissorType, checks, views, className }: Props) {
  const file = DIAGRAM_FILE[scissorType] || DIAGRAM_FILE['기타'];
  // 받은 그림을 파일명과 함께 보관 — 종류가 바뀌면(file 불일치) 로딩으로 간주 (effect 안 동기 setState 회피)
  const [loaded, setLoaded] = useState<{ file: string; src: DiagramSource | null } | null>(null);
  const uid = useId().replace(/[^a-zA-Z0-9]/g, '');

  useEffect(() => {
    let alive = true;
    loadDiagram(file).then((r) => { if (alive) setLoaded({ file, src: r }); });
    return () => { alive = false; };
  }, [file]);

  const src = loaded && loaded.file === file ? loaded.src : undefined;

  if (src === undefined) return <div className={`aspect-[1000/720] rounded bg-neutral-50 animate-pulse ${className || ''}`} />;
  if (src === null) return <p className={`text-xs text-neutral-400 ${className || ''}`}>{scissorType} 그림이 아직 없습니다 — 고객에게는 목록만 표시됩니다.</p>;

  const layers = woundLayers(checks);
  const show = views || woundViews(checks);
  return (
    <div className={`space-y-2 ${className || ''}`}>
      {show.map((v) => {
        const html = renderDiagramView(src, v, layers, `${uid}${v}`);
        return html ? (
          <div key={v} className="rounded border border-neutral-100 bg-white overflow-hidden [&>svg]:block [&>svg]:w-full [&>svg]:h-auto" dangerouslySetInnerHTML={{ __html: html }} />
        ) : null;
      })}
    </div>
  );
}
