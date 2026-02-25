'use client';

import { useRouter } from 'next/navigation';
import { cn } from '@/lib/utils/cn';

interface PipelineStage {
  label: string;
  count: number;
  color?: string;  // bg-xxx 클래스
  href?: string;   // 클릭 시 이동 경로
}

interface PipelineBarProps {
  stages: PipelineStage[];
  basePath?: string; // e.g. '/orders' → /orders?status=xxx
}

/** 주문/복원수리 상태 파이프라인 수평 바 */
export function PipelineBar({ stages, basePath }: PipelineBarProps) {
  const router = useRouter();
  const total = stages.reduce((sum, s) => sum + s.count, 0);

  // 기본 색상 순환
  const defaultColors = [
    'bg-info/80', 'bg-info/60', 'bg-terracotta/70',
    'bg-terracotta/50', 'bg-warning/60', 'bg-success/70',
    'bg-success/50', 'bg-neutral-300',
  ];

  return (
    <div className="space-y-2">
      {/* 바 */}
      {total > 0 && (
        <div className="flex h-3 rounded-full overflow-hidden bg-neutral-100">
          {stages.map((stage, i) => {
            if (stage.count === 0) return null;
            const pct = (stage.count / total) * 100;
            return (
              <div
                key={stage.label}
                className={cn(
                  'h-full transition-all',
                  stage.color || defaultColors[i % defaultColors.length],
                  stage.href || basePath ? 'cursor-pointer hover:opacity-80' : ''
                )}
                style={{ width: `${Math.max(pct, 3)}%` }}
                onClick={() => {
                  const target = stage.href || (basePath ? `${basePath}?status=${(stage as any).status || ''}` : '');
                  if (target) router.push(target);
                }}
                title={`${stage.label}: ${stage.count}건`}
              />
            );
          })}
        </div>
      )}

      {/* 범례 */}
      <div className="flex flex-wrap gap-x-4 gap-y-1 overflow-x-auto">
        {stages.map((stage, i) => (
          <button
            key={stage.label}
            type="button"
            className="flex items-center gap-1.5 text-xs text-neutral-600 hover:text-indigo-black transition whitespace-nowrap"
            onClick={() => {
              const target = stage.href || (basePath ? `${basePath}?status=${(stage as any).status || ''}` : '');
              if (target) router.push(target);
            }}
          >
            <span
              className={cn(
                'w-2.5 h-2.5 rounded-full shrink-0',
                stage.color || defaultColors[i % defaultColors.length]
              )}
            />
            {stage.label}
            <span className="font-bold text-indigo-black">{stage.count}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
