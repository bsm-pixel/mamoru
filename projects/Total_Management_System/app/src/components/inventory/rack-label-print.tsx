'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { Printer } from 'lucide-react';
import type { LocationWithProducts } from '@/hooks/use-warehouse';

/**
 * 렉 위치라벨 인쇄 — 115, 2026-07-18
 *
 * 자리 주소(R01-3-G3)를 실물 렉/수납함에 붙일 라벨로 뽑는다.
 * 로케이션 관리는 "시스템에만 있고 현장엔 없으면" 무용지물이라 이 인쇄가 사실상 필수 단계.
 *
 * 새 탭 window.open 패턴 — 기존 InventoryPrintModal / POPrintModal 과 동일.
 */

interface Props {
  rackNo: number;
  rackLabel: string | null;
  locations: LocationWithProducts[];   // 이 렉의 자리들 (정렬된 상태로 받음)
  onClose: () => void;
}

/** 라벨 크기 프리셋 — 실물 렉/수납함 칸 크기에 맞춰 고른다 */
const SIZES = {
  large:  { key: 'large',  label: '큰 라벨 (렉·선반용)',   cols: 3, codeSize: 26, nameSize: 11, pad: 12 },
  medium: { key: 'medium', label: '중간 (일반 칸)',        cols: 5, codeSize: 18, nameSize: 9,  pad: 8 },
  small:  { key: 'small',  label: '작은 라벨 (수납함 칸)',  cols: 8, codeSize: 13, nameSize: 0,  pad: 5 },
} as const;
type SizeKey = keyof typeof SIZES;

export function RackLabelPrint({ rackNo, rackLabel, locations, onClose }: Props) {
  const [size, setSize] = useState<SizeKey>('medium');
  const [includeProduct, setIncludeProduct] = useState(true);
  const [onlyAssigned, setOnlyAssigned] = useState(false);

  const rows = onlyAssigned ? locations.filter((l) => l.product_count > 0) : locations;
  const cfg = SIZES[size];

  const handlePrint = () => {
    const w = window.open('', '_blank');
    if (!w) return;

    const cells = rows.map((l) => {
      const productLine =
        includeProduct && cfg.nameSize > 0 && l.products.length > 0
          ? `<div class="nm">${escapeHtml(l.products[0].name)}${l.product_count > 1 ? ` 외 ${l.product_count - 1}` : ''}</div>`
          : '';
      const labelLine = cfg.nameSize > 0 ? `<div class="lb">${escapeHtml(l.label || '')}</div>` : '';
      return `<div class="cell"><div class="cd">${escapeHtml(l.code)}</div>${labelLine}${productLine}</div>`;
    }).join('');

    w.document.write(`
      <html><head><title>위치라벨 — ${rackNo}번 렉</title>
      <style>
        @page { margin: 10mm; }
        body { font-family: 'Noto Sans KR','Apple SD Gothic Neo',sans-serif; color:#000; margin:0; }
        h1 { font-size: 15px; font-weight: 800; margin: 0 0 2px; }
        .sub { font-size: 10px; color:#666; margin-bottom: 10px; }
        .grid { display: grid; grid-template-columns: repeat(${cfg.cols}, 1fr); gap: 4mm; }
        .cell {
          border: 1.5px solid #000; border-radius: 3px; padding: ${cfg.pad}px 6px;
          text-align: center; break-inside: avoid; page-break-inside: avoid;
          display: flex; flex-direction: column; justify-content: center; min-height: ${cfg.pad * 3 + 18}px;
        }
        .cd { font-family: 'Courier New', monospace; font-weight: 800; font-size: ${cfg.codeSize}px; letter-spacing: 1px; }
        .lb { font-size: ${cfg.nameSize}px; color:#444; margin-top: 3px; }
        .nm { font-size: ${cfg.nameSize}px; color:#000; font-weight:600; margin-top: 2px;
              overflow:hidden; white-space:nowrap; text-overflow:ellipsis; }
        .cut { font-size: 9px; color:#999; margin-top: 12px; }
      </style></head><body>
        <h1>${rackNo}번 렉${rackLabel ? ` · ${escapeHtml(rackLabel)}` : ''} 위치라벨</h1>
        <div class="sub">${rows.length}개 · 잘라서 해당 자리에 붙이세요 (1단 = 맨 아래)</div>
        <div class="grid">${cells}</div>
        <div class="cut">※ 코드가 실물과 일치해야 배치도 검색이 의미를 갖습니다.</div>
      </body></html>
    `);
    w.document.close();
    w.print();
  };

  return (
    <Modal
      open
      onClose={onClose}
      title={`${rackNo}번 렉 위치라벨 인쇄`}
      className="max-w-2xl"
      preventAutoClose
    >
      <div className="space-y-4">
        {/* 라벨 크기 */}
        <div>
          <label className="text-xs text-neutral-500 mb-1.5 block">라벨 크기</label>
          <div className="flex gap-1.5">
            {(Object.keys(SIZES) as SizeKey[]).map((k) => (
              <button
                key={k}
                onClick={() => setSize(k)}
                className={`flex-1 py-2 px-2 rounded-lg text-xs font-semibold transition border ${
                  size === k ? 'bg-neutral-900 text-white border-neutral-900'
                             : 'bg-white text-neutral-600 border-neutral-200 hover:border-neutral-400'
                }`}
              >
                {SIZES[k].label}
                <span className="block text-[10px] font-normal opacity-70">한 줄 {SIZES[k].cols}개</span>
              </button>
            ))}
          </div>
        </div>

        {/* 옵션 */}
        <div className="flex flex-wrap gap-4">
          <label className="flex items-center gap-1.5 text-xs text-neutral-600">
            <input type="checkbox" checked={includeProduct} onChange={(e) => setIncludeProduct(e.target.checked)}
              disabled={cfg.nameSize === 0} className="w-3.5 h-3.5" />
            제품명도 표시
            {cfg.nameSize === 0 && <span className="text-neutral-400">(작은 라벨은 코드만)</span>}
          </label>
          <label className="flex items-center gap-1.5 text-xs text-neutral-600">
            <input type="checkbox" checked={onlyAssigned} onChange={(e) => setOnlyAssigned(e.target.checked)}
              className="w-3.5 h-3.5" />
            제품 있는 자리만
          </label>
        </div>

        {/* 미리보기 */}
        <div>
          <p className="text-xs text-neutral-500 mb-1.5">미리보기 · 총 {rows.length}개</p>
          <div className="border border-neutral-200 rounded-lg p-3 bg-neutral-50 max-h-64 overflow-y-auto">
            {rows.length === 0 ? (
              <p className="text-xs text-neutral-400 text-center py-6">인쇄할 자리가 없습니다</p>
            ) : (
              <div className="grid gap-2" style={{ gridTemplateColumns: `repeat(${Math.min(cfg.cols, 6)}, 1fr)` }}>
                {rows.slice(0, 24).map((l) => (
                  <div key={l.id} className="border border-neutral-800 rounded bg-white px-1.5 py-2 text-center">
                    <div className="font-mono font-bold text-indigo-black leading-tight"
                      style={{ fontSize: Math.min(cfg.codeSize, 15) }}>{l.code}</div>
                    {cfg.nameSize > 0 && (
                      <div className="text-[9px] text-neutral-500 truncate">{l.label}</div>
                    )}
                    {includeProduct && cfg.nameSize > 0 && l.products[0] && (
                      <div className="text-[9px] font-semibold text-indigo-black truncate">{l.products[0].name}</div>
                    )}
                  </div>
                ))}
              </div>
            )}
            {rows.length > 24 && (
              <p className="text-[11px] text-neutral-400 mt-2 text-center">…외 {rows.length - 24}개 (인쇄엔 전부 포함)</p>
            )}
          </div>
        </div>

        <div className="flex gap-2 pt-1">
          <Button variant="ghost" className="flex-1" onClick={onClose}>닫기</Button>
          <Button className="flex-1" onClick={handlePrint} disabled={rows.length === 0}>
            <Printer size={14} /> 인쇄
          </Button>
        </div>
      </div>
    </Modal>
  );
}

function escapeHtml(s: string): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
