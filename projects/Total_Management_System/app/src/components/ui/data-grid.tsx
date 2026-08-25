'use client';

import { memo } from 'react';
import { LayoutGrid } from 'lucide-react';
import { cn } from '@/lib/utils/cn';
import { Button } from './button';

/**
 * DataGrid — 마스터-디테일 전용 밀집 표 (2026-07-16)
 *
 * 판매관리 SalesGridTable에서 검증된 밀집표(sticky 헤더 + 선택 하이라이트 + 촘촘한 패딩)를
 * 도메인 무관하게 일반화. 컬럼 정의·행 렌더·행 색줄은 각 페이지가 주입한다.
 * (마스터-디테일 밀집표 — sticky 헤더 + 선택 하이라이트 + 촘촘한 패딩)
 */
export interface GridColumn<T> {
  key: string;
  label: string;
  align?: 'left' | 'right' | 'center';
  headClassName?: string;
  cellClassName?: string;
  render: (row: T) => React.ReactNode;
}

interface DataGridProps<T> {
  columns: GridColumn<T>[];
  rows: T[];
  getRowKey: (row: T) => string;
  /** 현재 선택된 행 key — 좌측 inset 하이라이트 */
  selectedKey?: string;
  onSelect?: (row: T) => void;
  /** 행별 추가 클래스(취소 opacity, 좌측 색줄 등) */
  rowClassName?: (row: T) => string;
  emptyMessage?: string;
}

const alignCls = (a?: string) => (a === 'right' ? 'text-right' : a === 'center' ? 'text-center' : 'text-left');

function DataGridInner<T>({
  columns,
  rows,
  getRowKey,
  selectedKey,
  onSelect,
  rowClassName,
  emptyMessage = '데이터가 없습니다',
}: DataGridProps<T>) {
  if (rows.length === 0) {
    return <div className="flex items-center justify-center h-40 text-sm text-neutral-400">{emptyMessage}</div>;
  }
  return (
    <table className="w-full text-sm border-collapse">
      <thead>
        <tr className="sticky top-0 bg-stone-50 z-[1] text-left text-[11px] font-semibold text-neutral-500">
          {columns.map((c) => (
            <th key={c.key} className={cn('px-3 py-2.5 whitespace-nowrap', alignCls(c.align), c.headClassName)}>
              {c.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((row) => {
          const key = getRowKey(row);
          const sel = selectedKey === key;
          return (
            <tr
              key={key}
              onClick={() => onSelect?.(row)}
              className={cn(
                'border-b border-neutral-100 transition',
                onSelect && 'cursor-pointer',
                sel ? 'bg-[#F4F0EA] shadow-[inset_3px_0_0_#1A1A1A]' : 'hover:bg-warm-ivory/50',
                rowClassName?.(row),
              )}
            >
              {columns.map((c) => (
                <td key={c.key} className={cn('px-3 py-2.5', alignCls(c.align), c.cellClassName)}>
                  {c.render(row)}
                </td>
              ))}
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

// memo로 감싸되 제네릭 시그니처 보존
export const DataGrid = memo(DataGridInner) as typeof DataGridInner;

/** PC 그리드 토글 버튼 — isLg일 때만 노출 (Topbar action에 배치) */
export function GridToggleButton({
  isLg,
  gridMode,
  onToggle,
}: {
  isLg: boolean;
  gridMode: boolean;
  onToggle: () => void;
}) {
  if (!isLg) return null;
  return (
    <Button onClick={onToggle} size="sm" variant="secondary" title="목록을 밀집 표로 보기">
      <LayoutGrid size={14} />
      {gridMode ? '카드 보기' : 'PC 그리드'}
    </Button>
  );
}
