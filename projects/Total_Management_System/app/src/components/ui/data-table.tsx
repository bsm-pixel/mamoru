'use client';

import { cn } from '@/lib/utils/cn';
import { Skeleton } from './skeleton';

interface Column<T> {
  key: string;
  label: string;
  className?: string;
  render?: (row: T) => React.ReactNode;
}

interface DataTableProps<T> {
  columns: Column<T>[];
  data: T[];
  loading?: boolean;
  onRowClick?: (row: T) => void;
  emptyMessage?: string;
  keyExtractor: (row: T) => string;
}

export function DataTable<T extends Record<string, unknown>>({
  columns,
  data,
  loading,
  onRowClick,
  emptyMessage = '데이터가 없습니다',
  keyExtractor,
}: DataTableProps<T>) {
  if (loading) {
    return (
      <div className="space-y-3">
        {Array.from({ length: 5 }).map((_, i) => (
          <Skeleton key={i} className="h-14 w-full" />
        ))}
      </div>
    );
  }

  if (data.length === 0) {
    return (
      <div className="flex items-center justify-center h-40 text-sm text-neutral-400">
        {emptyMessage}
      </div>
    );
  }

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-neutral-200">
            {columns.map((col) => (
              <th
                key={col.key}
                className={cn(
                  'text-left py-3 px-3 text-xs font-semibold text-neutral-500 uppercase tracking-wide',
                  col.className
                )}
              >
                {col.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.map((row) => (
            <tr
              key={keyExtractor(row)}
              onClick={() => onRowClick?.(row)}
              className={cn(
                'border-b border-neutral-100 last:border-0 transition',
                onRowClick && 'cursor-pointer hover:bg-warm-ivory/60'
              )}
            >
              {columns.map((col) => (
                <td key={col.key} className={cn('py-3 px-3', col.className)}>
                  {col.render
                    ? col.render(row)
                    : String(row[col.key] ?? '')}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
