import { cn } from '@/lib/utils/cn';

interface CardProps {
  children: React.ReactNode;
  className?: string;
  padding?: boolean;
  onClick?: () => void;
  /** 목록 키보드 표시(data-list-row · data-selected) 같은 data-* 는 그대로 넘긴다 */
  [key: `data-${string}`]: string | undefined;
}

export function Card({ children, className, padding = true, onClick, ...data }: CardProps) {
  return (
    <div
      {...data}
      onClick={onClick}
      className={cn(
        'bg-card-white rounded-xl border border-neutral-200 shadow-sm',
        padding && 'p-5',
        className
      )}
    >
      {children}
    </div>
  );
}

export function CardHeader({ children, className }: { children: React.ReactNode; className?: string }) {
  return (
    <div className={cn('flex items-center justify-between mb-4', className)}>
      {children}
    </div>
  );
}

export function CardTitle({ children, className }: { children: React.ReactNode; className?: string }) {
  return (
    <h3 className={cn('text-base font-bold text-indigo-black', className)}>
      {children}
    </h3>
  );
}
