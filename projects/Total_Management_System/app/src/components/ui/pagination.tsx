import { Button } from './button';
import { ChevronLeft, ChevronRight } from 'lucide-react';

interface PaginationProps {
  page: number;
  totalPages: number;
  total: number;
  onPageChange: (page: number) => void;
  /** 건수 단위 (기본: '건') */
  unit?: string;
}

/** 목록 페이지 공통 페이지네이션 (건수 + 이전/다음) */
export function Pagination({ page, totalPages, total, onPageChange, unit = '건' }: PaginationProps) {
  return (
    <div className="flex items-center justify-between">
      <span className="text-xs text-neutral-400">총 {total}{unit}</span>
      {totalPages > 1 && (
        <div className="flex items-center gap-2">
          <Button variant="ghost" size="sm" disabled={page <= 1} onClick={() => onPageChange(page - 1)}>
            <ChevronLeft size={16} />
          </Button>
          <span className="text-sm text-neutral-500">{page} / {totalPages}</span>
          <Button variant="ghost" size="sm" disabled={page >= totalPages} onClick={() => onPageChange(page + 1)}>
            <ChevronRight size={16} />
          </Button>
        </div>
      )}
    </div>
  );
}
