'use client';

import { Modal } from '@/components/ui/modal';
import { Trash2, Plus, AlertTriangle } from 'lucide-react';
import { useAddCol, useAddRow, useAddLevel, useDeleteLocation, type LocationWithProducts } from '@/hooks/use-warehouse';
import { cellGridPos } from '@/lib/warehouse/location-code';

/**
 * 렉 구조 수정 모달 — 115, 2026-07-18
 *
 * 배치도 진입 화면은 '보는 것'에 집중시키고(깔끔하게), 구조 편집은 전부 여기로 모은다.
 * 편집 중 실수로 닫히면 안 되므로 preventAutoClose (Escape·배경 드래그로 안 닫힘, X 로만 닫기).
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

export function RackEditModal({ rackNo, rackLabel, levels, onClose, onDeleteRack }: Props) {
  const addCol = useAddCol();
  const addRow = useAddRow();
  const addLevel = useAddLevel();
  const deleteLocation = useDeleteLocation();

  const totalCells = levels.reduce((s, l) => s + l.cells.length, 0);
  const assignedCells = levels.reduce((s, l) => s + l.cells.filter((c) => c.product_count > 0).length, 0);
  const busy = addCol.isPending || addRow.isPending || addLevel.isPending || deleteLocation.isPending;

  return (
    <Modal
      open
      onClose={onClose}
      title={`${rackNo}번 렉 수정${rackLabel ? ` · ${rackLabel}` : ''}`}
      className="max-w-2xl"
      preventAutoClose
    >
      <div className="space-y-4">
        <div className="flex items-center justify-between">
          <p className="text-xs text-neutral-500">
            자리 {totalCells}개 · 제품 있는 자리 {assignedCells}개
            <span className="block text-[11px] text-neutral-400 mt-0.5">1단 = 맨 아래 · 아래 목록은 실제 렉처럼 위가 큰 번호</span>
          </p>
          <button
            onClick={() => addLevel.mutate({ rack_no: rackNo, cols: 0 })}
            disabled={busy}
            className="text-xs font-semibold px-3 py-1.5 rounded-lg border border-neutral-200 hover:bg-neutral-50 transition disabled:opacity-50"
            title="맨 위에 단을 하나 더 만듭니다 (칸 없는 선반으로 생성)"
          >
            <Plus size={12} className="inline" /> 맨 위에 단 추가
          </button>
        </div>

        {/* 단별 편집 */}
        <div className="space-y-2 max-h-[46vh] overflow-y-auto pr-1">
          {levels.map((lvl) => {
            const shape = lvl.isShelf ? '선반 (칸 없음)'
              : lvl.isDrawer ? `수납함 ${lvl.rows}행 ${lvl.cols}열 = ${lvl.cells.length}칸`
              : `${lvl.cols}칸 (한 줄)`;
            const used = lvl.cells.filter((c) => c.product_count > 0).length;
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
                      disabled={busy}
                      className="text-[11px] font-semibold px-2 py-1 rounded border border-neutral-200 hover:bg-neutral-50 disabled:opacity-50"
                      title="이 단에 열(칸) 하나 추가"
                    >＋열</button>
                    <button
                      onClick={() => addRow.mutate({ rack_no: rackNo, level_no: lvl.levelNo })}
                      disabled={busy}
                      className="text-[11px] font-semibold px-2 py-1 rounded border border-neutral-200 hover:bg-neutral-50 disabled:opacity-50"
                      title="이 단에 행 추가 (수납함으로 만들기)"
                    >＋행</button>
                  </div>
                </div>

                {/* 칸 목록 — 실제 배치 그대로(열×행). 클릭하면 그 칸만 삭제 */}
                {!lvl.isShelf && (
                  <div
                    className="grid gap-1 mt-2 overflow-x-auto pb-1"  /* 열이 아주 많은 수납함은 가로 스크롤 */
                    style={{ gridTemplateColumns: `repeat(${Math.max(1, lvl.cols)}, 26px)`, justifyContent: 'start' }}
                  >
                    {lvl.cells.map((c) => {
                      // 중간 칸을 지워도 나머지가 밀리지 않게 열·행 명시 배치 (배치도·인쇄와 동일 규칙)
                      const pos = cellGridPos(c.bin_no, c.bin_row);
                      return (
                        <button
                          key={c.id}
                          onClick={() => {
                            if (c.product_count > 0) {
                              if (!window.confirm(`'${c.code}' 에 제품 ${c.product_count}종이 있습니다.\n삭제하면 그 제품들은 '위치 미지정'이 됩니다. 계속할까요?`)) return;
                            }
                            deleteLocation.mutate({ id: c.id });
                          }}
                          disabled={busy}
                          style={{ gridColumn: pos.col, gridRow: pos.row }}
                          title={`${c.label}${c.product_count > 0 ? ` · 제품 ${c.product_count}종` : ''} — 클릭하면 이 칸 삭제`}
                          className={`w-[26px] h-[26px] flex items-center justify-center text-[9px] font-mono rounded border transition disabled:opacity-50 ${
                            c.product_count > 0
                              ? 'border-emerald-300 bg-emerald-50 text-emerald-700 hover:border-red-400 hover:bg-red-50 hover:text-red-600'
                              : 'border-neutral-200 text-neutral-400 hover:border-red-400 hover:text-red-600'
                          }`}
                        >
                          {c.code.split('-').slice(2).join('-') || c.code}
                        </button>
                      );
                    })}
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
