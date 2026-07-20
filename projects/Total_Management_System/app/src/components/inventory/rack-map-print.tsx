'use client';

import { useState, useMemo } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { Printer, AlertTriangle } from 'lucide-react';
import { buildRackPrintHtml, fitCellH, A4, type PrintRack, type Orient } from '@/lib/warehouse/rack-print-html';

/**
 * 렉 배치도 인쇄 모달 — A4 한 장에 렉 하나 (2026-07-20)
 * 인쇄물 HTML 자체는 lib/warehouse/rack-print-html.ts (순수 함수·헤드리스 검증용).
 * 여기서는 옵션 UI + 미리보기 + window.open 인쇄만 담당한다.
 */

interface Props {
  racks: PrintRack[];        // 전체 렉 (모든 렉 인쇄용)
  targetRackNo: number;      // 인쇄 버튼을 누른 렉
  onClose: () => void;
}

export function RackMapPrint({ racks, targetRackNo, onClose }: Props) {
  const [orient, setOrient] = useState<Orient>('portrait');
  const [allRacks, setAllRacks] = useState(false);
  const [showProduct, setShowProduct] = useState(true);

  const targets = useMemo(
    () => (allRacks ? racks : racks.filter((r) => r.rackNo === targetRackNo)),
    [allRacks, racks, targetRackNo],
  );

  // 자리가 너무 많아 글씨가 작아질 렉 — 가로 방향을 권한다
  const tight = targets.filter((r) => fitCellH(r, orient) <= 7).map((r) => r.rackNo);

  const html = useMemo(
    () => buildRackPrintHtml(targets, { orient, showProduct }),
    [targets, orient, showProduct],
  );

  const handlePrint = () => {
    const w = window.open('', '_blank');
    if (!w) return;
    w.document.write(html);
    w.document.close();
    w.focus();
    w.print();
  };

  const px = (mm: number) => mm * 3.7795;               // 96dpi 기준 mm→px
  const previewW = 300;
  const scale = previewW / px(A4[orient].w);
  const previewH = Math.min(px(A4[orient].h) * scale * targets.length, 420);

  return (
    <Modal
      open
      onClose={onClose}
      title={`${targetRackNo}번 렉 배치도 인쇄`}
      className="max-w-3xl"
      preventAutoClose
    >
      <div className="flex flex-col sm:flex-row gap-5">
        {/* 설정 */}
        <div className="flex-1 min-w-0 space-y-4">
          <p className="text-xs text-neutral-500">
            A4 <strong className="text-indigo-black">한 장에 렉 하나</strong>가 실물 모양 그대로 인쇄됩니다.
            <span className="block text-[11px] text-neutral-400 mt-0.5">렉 옆면에 붙여 두고 눈으로 대조하는 용도 (라벨지 아님)</span>
          </p>

          <div>
            <label className="text-xs text-neutral-500 mb-1.5 block">용지 방향</label>
            <div className="flex gap-1.5">
              {([['portrait', '세로', '단이 많을 때'], ['landscape', '가로', '칸(열)이 많을 때']] as const).map(([k, name, hint]) => (
                <button
                  key={k}
                  onClick={() => setOrient(k)}
                  className={`flex-1 py-2 rounded-lg text-xs font-semibold transition border ${
                    orient === k ? 'bg-neutral-900 text-white border-neutral-900'
                                 : 'bg-white text-neutral-600 border-neutral-200 hover:border-neutral-400'
                  }`}
                >
                  {name}
                  <span className="block text-[10px] font-normal opacity-70">{hint}</span>
                </button>
              ))}
            </div>
          </div>

          <div className="space-y-2">
            <label className="flex items-center gap-1.5 text-xs text-neutral-600">
              <input type="checkbox" checked={showProduct} onChange={(e) => setShowProduct(e.target.checked)} className="w-3.5 h-3.5" />
              제품명·재고 표시 <span className="text-neutral-400">(끄면 자리 주소만)</span>
            </label>
            <label className="flex items-center gap-1.5 text-xs text-neutral-600">
              <input type="checkbox" checked={allRacks} onChange={(e) => setAllRacks(e.target.checked)}
                disabled={racks.length < 2} className="w-3.5 h-3.5" />
              모든 렉 인쇄 <span className="text-neutral-400">(렉마다 한 장 · 총 {racks.length}장)</span>
            </label>
          </div>

          {tight.length > 0 && (
            <p className="flex items-start gap-1.5 text-[11px] text-amber-600">
              <AlertTriangle size={12} className="mt-0.5 shrink-0" />
              {tight.join('·')}번 렉은 자리가 많아 글씨가 작아집니다.
              {orient === 'portrait' && ' 가로 방향을 권합니다.'}
            </p>
          )}

          <div className="flex gap-2 pt-1">
            <Button variant="ghost" className="flex-1" onClick={onClose}>닫기</Button>
            <Button className="flex-1" onClick={handlePrint} disabled={targets.length === 0}>
              <Printer size={14} /> 인쇄 ({targets.length}장)
            </Button>
          </div>
        </div>

        {/* 미리보기 — 인쇄와 '같은 HTML' 을 그대로 축소해 보여준다 (실물과 어긋날 여지 0) */}
        <div className="shrink-0" style={{ width: previewW }}>
          <p className="text-[11px] text-neutral-400 mb-1.5">
            미리보기 · A4 {orient === 'portrait' ? '세로' : '가로'} {targets.length}장
          </p>
          <div className="border border-neutral-200 rounded-lg overflow-hidden bg-neutral-100"
            style={{ width: previewW, height: previewH, overflowY: 'auto' }}>
            <div style={{ width: px(A4[orient].w), transform: `scale(${scale})`, transformOrigin: 'top left' }}>
              <iframe
                title="인쇄 미리보기"
                srcDoc={html}
                scrolling="no"
                style={{ width: px(A4[orient].w), height: px(A4[orient].h) * targets.length, border: 0, display: 'block' }}
              />
            </div>
          </div>
        </div>
      </div>
    </Modal>
  );
}
