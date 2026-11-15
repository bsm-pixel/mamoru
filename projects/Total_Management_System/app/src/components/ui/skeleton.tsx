import { cn } from '@/lib/utils/cn';

export function Skeleton({ className }: { className?: string }) {
  return (
    <div
      className={cn(
        'animate-pulse rounded-lg bg-neutral-200/60',
        className
      )}
    />
  );
}

/**
 * 불러오는 동안 보이는 회색 뼈대 줄 (2026-11-15 통일)
 * 전엔 화면마다 "로딩중..." · "불러오는 중…" 글자와 뼈대가 섞여 있었다 → 목록·칸을 불러올 때는 이것 하나.
 * 화면 읽기 도구에는 "불러오는 중"으로 읽힌다.
 */
export function SkeletonRows({ rows = 4, className }: { rows?: number; className?: string }) {
  const widths = ['w-11/12', 'w-4/5', 'w-full', 'w-3/4', 'w-5/6', 'w-2/3'];
  return (
    <div role="status" aria-label="불러오는 중" className={cn('space-y-2.5 p-4', className)}>
      {Array.from({ length: rows }).map((_, i) => (
        <Skeleton key={i} className={cn('h-4', widths[i % widths.length])} />
      ))}
    </div>
  );
}
