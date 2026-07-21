'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Trash2, Plus, AlertTriangle, Combine, X, Scissors } from 'lucide-react';
import {
  useAddCol, useAddRow, useAddLevel, useAddCell, useDeleteLocation,
  useMergeCells, useSplitCell, type LocationWithProducts,
} from '@/hooks/use-warehouse';
import { cellSpan, binLetter } from '@/lib/warehouse/location-code';

/**
 * 렉 구조 수정 모달 — 115 / 116, 2026-07-18~20
 *
 * 배치도 진입 화면은 '보는 것'에 집중시키고(깔끔하게), 구조 편집은 전부 여기로 모은다.
 * 편집 중 실수로 닫히면 안 되므로 preventAutoClose (Escape·배경 드래그로 안 닫힘, X 로만 닫기).
 *
 * 칸 조작(116):
 *  · 일반 모드 — 칸 클릭 = 삭제(점선 빈자리로 남음) / 점선 빈자리 클릭 = 다시 생성
 *  · 병합 모드 — 칸 클릭 = 선택, 같은 행 인접 2칸 이상 고르면 '가로로 합치기'
 *  · 넓은 칸(병합됨)은 ✂ 로 다시 분리
 */

export interface RackLevelView {
  levelNo: number;
  cells: LocationWithProducts[];
  rows: number;
  cols: number;
  isShelf: boolean;
  isDrawer: boolean;
}

interface Props {
  rackNo: number;
  rackLabel: string | null;
  levels: RackLevelView[];         // 화면과 동일하게 내림차순(위→아래)으로 받는다
  onClose: () => void;
  onDeleteRack: () => void;
}

/** 코드 뒷자리 (R01-3-C1 → C1). 병합 칸이면 범위로 (C1~D1) */
function cellTail(c: LocationWithProducts): string {
  const tail = c.code.split('-').slice(2).join('-') || c.code;
  const span = cellSpan(c.col_span);
  if (span <= 1) return tail;
  const col = c.bin_no ?? 1;
  const row = c.bin_row ?? 1;
  const rows = Math.max(1, row); // 행 표기 여부는 원 코드가 이미 반영
  const hasRowSuffix = /\d$/.test(tail) && rows > 0;
  const endLetter = binLetter(col + span - 1);
  const end = hasRowSuffix ? `${endLetter}${row}` : endLetter;
  return `${tail}~${end}`;
}

export function RackEditModal({ rackNo, rackLabel, levels, onClose, onDeleteRack }: Props) {
  const addCol = useAddCol();
  const addRow = useAddRow();
  const addLevel = useAddLevel();
  const addCell = useAddCell();
  const deleteLocation = useDeleteLocation();
  const mergeCells = useMergeCells();
  const splitCell = useSplitCell();

  const [mergeMode, setMergeMode] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  const totalCells = levels.reduce((s, l) => s + l.cells.length, 0);
  const assignedCells = levels.reduce((s, l) => s + l.cells.filter((c) => c.product_count > 0).length, 0);
  const busy = addCol.isPending || addRow.isPending || addLevel.isPending || addCell.isPending
    || deleteLocation.isPending || mergeCells.isPending || splitCell.isPending;

  const clearSel = () => setSelected(new Set());
  const toggleSel = (id: string) => setSelected((prev) => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  const doMerge = () => {
    if (selected.size < 2) return;
    mergeCells.mutate({ ids: [...selected] });
    clearSel();
  };

  return (
    <Modal
      open
      onClose={onClose}
      title={`${rackNo}번 렉 수정${rackLabel ? ` · ${rackLabel}` : ''}`}
      className="max-w-2xl"
      preventAutoClose
    >
      <div className="space-y-4">
        <div className="flex items-start justify-between gap-2">
          <p className="text-xs text-neutral-500">
            자리 {totalCells}개 · 제품 있는 자리 {assignedCells}개
            <span className="block text-[11px] text-neutral-400 mt-0.5">
              1단 = 맨 아래 · {mergeMode ? '합칠 칸들을 고르세요 (같은 행·서로 붙은 칸)' : '칸 클릭 = 삭제 / 점선 빈자리 클릭 = 생성'}
            </span>
          </p>
          <div className="flex gap-1.5 shrink-0">
            <button
              onClick={() => { setMergeMode((v) => !v); clearSel(); }}
              disabled={busy}
              className={`flex items-center gap-1 text-xs font-semibold px-3 py-1.5 rounded-lg border transition disabled:opacity-50 ${
                mergeMode ? 'bg-neutral-900 text-white border-neutral-900' : 'border-neutral-200 hover:bg-neutral-50'
              }`}
              title="가로로 긴 칸막이 만들기 — 칸 여러 개를 하나로"
            >
              <Combine size={12} /> 병합 {mergeMode ? 'ON' : ''}
            </button>
            <button
              onClick={() => addLevel.mutate({ rack_no: rackNo, cols: 0 })}
              disabled={busy || mergeMode}
              className="text-xs font-semibold px-3 py-1.5 rounded-lg border border-neutral-200 hover:bg-neutral-50 transition disabled:opacity-40"
              title="맨 위에 단을 하나 더 만듭니다 (칸 없는 선반으로 생성)"
            >
              <Plus size={12} className="inline" /> 맨 위에 단
            </button>
          </div>
        </div>

        {/* 병합 실행 바 */}
        {mergeMode && (
          <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-neutral-900 text-white">
            <span className="text-xs">{selected.size === 0 ? '합칠 칸을 고르세요' : `${selected.size}칸 선택됨`}</span>
            <div className="ml-auto flex gap-1.5">
              {selected.size > 0 && (
                <button onClick={clearSel} className="flex items-center gap-1 text-[11px] px-2 py-1 rounded bg-white/15 hover:bg-white/25 transition">
                  <X size={11} /> 선택 해제
                </button>
              )}
              <button
                onClick={doMerge}
                disabled={selected.size < 2 || busy}
                className="flex items-center gap-1 text-[11px] font-semibold px-2.5 py-1 rounded bg-white text-neutral-900 disabled:opacity-40"
              >
                <Combine size={11} /> {selected.size >= 2 ? `${selected.size}칸 가로로 합치기` : '2칸 이상 필요'}
              </button>
            </div>
          </div>
        )}

        {/* 단별 편집 */}
        <div className="space-y-2 max-h-[46vh] overflow-y-auto pr-1">
          {levels.map((lvl) => {
            const shape = lvl.isShelf ? '선반 (칸 없음)'
              : lvl.isDrawer ? `수납함 ${lvl.rows}행 ${lvl.cols}열`
              : `${lvl.cols}칸 (한 줄)`;
            const used = lvl.cells.filter((c) => c.product_count > 0).length;

            // 116: 시작 위치·덮인 위치 계산 (병합 칸이 오른쪽 열들을 덮는다)
            const start = new Map<string, LocationWithProducts>();
            const covered = new Set<string>();
            for (const c of lvl.cells) {
              const col = c.bin_no ?? 1, row = c.bin_row ?? 1, span = cellSpan(c.col_span);
              start.set(`${row}:${col}`, c);
              for (let k = 1; k < span; k++) covered.add(`${row}:${col + k}`);
            }
            const rows = Math.max(1, lvl.rows);
            const cols = Math.max(1, lvl.cols);

            return (
              <div key={lvl.levelNo} className="border border-neutral-200 rounded-lg px-3 py-2.5">
                <div className="flex items-center justify-between gap-2">
                  <div className="min-w-0">
                    <span className="text-sm font-bold text-indigo-black">{lvl.levelNo}단</span>
                    <span className="ml-2 text-xs text-neutral-500">{shape}</span>
                    {used > 0 && <span className="ml-2 text-[11px] text-emerald-600">제품 {used}자리</span>}
                  </div>
                  <div className="flex gap-1.5 shrink-0">
                    <button
                      onClick={() => addCol.mutate({ rack_no: rackNo, level_no: lvl.levelNo })}
                      disabled={busy || mergeMode}
                      className="text-[11px] font-semibold px-2 py-1 rounded border border-neutral-200 hover:bg-neutral-50 disabled:opacity-40"
                      title="이 단에 열(칸) 하나 추가"
                    >＋열</button>
                    <button
                      onClick={() => addRow.mutate({ rack_no: rackNo, level_no: lvl.levelNo })}
                      disabled={busy || mergeMode}
                      className="text-[11px] font-semibold px-2 py-1 rounded border border-neutral-200 hover:bg-neutral-50 disabled:opacity-40"
                      title="이 단에 행 추가 (수납함으로 만들기)"
                    >＋행</button>
                  </div>
                </div>

                {/* 칸 격자 — 실제 배치 그대로. 빈자리는 점선, 있는 칸은 실선 */}
                {!lvl.isShelf && (
                  <div
                    className="grid gap-1 mt-2 overflow-x-auto pb-1"
                    style={{ gridTemplateColumns: `repeat(${cols}, 26px)`, justifyContent: 'start' }}
                  >
                    {Array.from({ length: rows }).flatMap((_, ri) =>
                      Array.from({ length: cols }).map((__, ci) => {
                        const row = ri + 1, col = ci + 1, key = `${row}:${col}`;
                        if (covered.has(key)) return null;         // 넓은 칸이 덮은 안쪽 열 — 아무것도 안 그림
                        const cell = start.get(key);
                        const span = cell ? cellSpan(cell.col_span) : 1;

                        if (!cell) {
                          // 점선 빈자리 — 클릭하면 생성 (병합 모드에선 비활성)
                          return (
                            <button
                              key={key}
                              onClick={() => !mergeMode && addCell.mutate({ rack_no: rackNo, level_no: lvl.levelNo, bin_no: col, bin_row: row })}
                              disabled={busy || mergeMode}
                              style={{ gridColumn: col, gridRow: row }}
                              title={mergeMode ? '' : `${binLetter(col)}${rows > 1 ? row : ''}칸 만들기`}
                              className={`w-full h-[26px] rounded border border-dashed border-neutral-200 text-neutral-300 text-[11px] transition disabled:opacity-50 ${
                                mergeMode ? '' : 'hover:border-neutral-500 hover:text-neutral-500'
                              }`}
                            >{mergeMode ? '' : '＋'}</button>
                          );
                        }

                        const sel = selected.has(cell.id);
                        const hasProd = cell.product_count > 0;
                        return (
                          <button
                            key={cell.id}
                            onClick={() => {
                              if (mergeMode) { toggleSel(cell.id); return; }
                              if (hasProd && !window.confirm(`'${cell.code}' 에 제품 ${cell.product_count}종이 있습니다.\n삭제하면 그 제품들은 '위치 미지정'이 됩니다. 계속할까요?`)) return;
                              deleteLocation.mutate({ id: cell.id });
                            }}
                            disabled={busy}
                            style={{ gridColumn: `${col} / span ${span}`, gridRow: row }}
                            title={
                              mergeMode ? `${cell.label} — 클릭하면 선택`
                                : `${cell.label}${hasProd ? ` · 제품 ${cell.product_count}종` : ''} — 클릭하면 삭제`
                            }
                            className={`relative h-[26px] px-1 flex items-center justify-center text-[9px] font-mono rounded border transition disabled:opacity-50 ${
                              sel ? 'border-neutral-900 ring-2 ring-neutral-900 bg-[#F4F0EA] text-indigo-black'
                              : hasProd ? 'border-emerald-300 bg-emerald-50 text-emerald-700 ' + (mergeMode ? 'hover:ring-1 hover:ring-neutral-400' : 'hover:border-red-400 hover:bg-red-50 hover:text-red-600')
                              : 'border-neutral-200 text-neutral-400 ' + (mergeMode ? 'hover:ring-1 hover:ring-neutral-400' : 'hover:border-red-400 hover:text-red-600')
                            }`}
                          >
                            <span className="truncate">{cellTail(cell)}</span>
                            {/* 넓은 칸 분리 — 일반 모드에서만 */}
                            {span > 1 && !mergeMode && (
                              <span
                                role="button"
                                tabIndex={0}
                                onClick={(e) => { e.stopPropagation(); if (!busy) splitCell.mutate({ id: cell.id }); }}
                                onKeyDown={(e) => { if (e.key === 'Enter') { e.stopPropagation(); if (!busy) splitCell.mutate({ id: cell.id }); } }}
                                title="가로 병합 해제"
                                className="absolute -top-1.5 -right-1.5 w-4 h-4 flex items-center justify-center rounded-full bg-neutral-800 text-white hover:bg-neutral-900"
                              >
                                <Scissors size={9} />
                              </span>
                            )}
                          </button>
                        );
                      }),
                    )}
                  </div>
                )}
                {lvl.isShelf && (
                  <p className="text-[11px] text-neutral-400 mt-1.5">
                    ＋열 을 누르면 칸으로 나눌 수 있습니다 (제품이 있으면 먼저 빼야 합니다)
                  </p>
                )}
              </div>
            );
          })}
        </div>

        {/* 렉 삭제 */}
        <div className="pt-3 border-t border-neutral-100">
          <button
            onClick={() => {
              const msg = assignedCells > 0
                ? `${rackNo}번 렉을 통째로 삭제합니다.\n제품이 있는 자리 ${assignedCells}개의 제품은 '위치 미지정'으로 바뀝니다.\n\n계속할까요?`
                : `${rackNo}번 렉을 통째로 삭제합니다. (자리 ${totalCells}개)\n\n계속할까요?`;
              if (!window.confirm(msg)) return;
              onDeleteRack();
            }}
            disabled={busy}
            className="w-full flex items-center justify-center gap-1.5 py-2 rounded-lg border border-red-200 text-xs font-semibold text-red-600 hover:bg-red-50 transition disabled:opacity-50"
          >
            <Trash2 size={13} /> 이 렉 통째로 삭제
          </button>
          {assignedCells > 0 && (
            <p className="flex items-start gap-1 text-[11px] text-amber-600 mt-1.5">
              <AlertTriangle size={12} className="mt-0.5 shrink-0" />
              제품이 배정된 자리가 있습니다. 삭제해도 <strong>제품·재고는 사라지지 않고</strong> 위치만 해제됩니다.
            </p>
          )}
        </div>
      </div>
    </Modal>
  );
}
