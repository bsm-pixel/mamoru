'use client';

import { useEffect, useMemo, useState } from 'react';
import { Button } from '@/components/ui/button';
import { X, Printer, FileDown, Minus, Plus } from 'lucide-react';
import toast from 'react-hot-toast';
import { buildRepairTagItems, renderRepairTagsPreview, renderRepairTagsZpl } from '@/lib/label/repair-tag';
import { sendZplToPrinter, downloadZpl } from '@/lib/label/browser-print';
import { backdropClose } from '@/lib/ui/backdrop';
import { EscClose } from '@/components/ui/esc-close';

interface Props {
  name: string;
  asId: string;
  /** 기본 수량 — 수리내역서가 있으면 그 가위 수, 없으면 접수 수량(마모루+타사) */
  defaultCount: number;
  onClose: () => void;
}

/** 복원수리 꼬리표 출력 — 수량 확인 후 한 번에 출력 (유포 꼬리표 라벨지 + 리본) */
export function RepairTagPrintModal({ name, asId, defaultCount, onClose }: Props) {
  const [count, setCount] = useState(Math.max(1, defaultCount || 1));
  const [previews, setPreviews] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const items = useMemo(() => buildRepairTagItems(name, asId, count), [name, asId, count]);

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const p = await renderRepairTagsPreview(items);
        if (alive) setPreviews(p);
      } catch { /* 미리보기 실패는 출력과 무관 */ }
    })();
    return () => { alive = false; };
  }, [items]);

  async function handlePrint() {
    setBusy(true);
    try {
      const zpl = await renderRepairTagsZpl(items);
      if (await sendZplToPrinter(zpl)) {
        toast.success(`꼬리표 ${count}장 출력 전송됨`);
        onClose();
      } else {
        toast.error('프린터(Browser Print) 미연결 — 프린터·Browser Print 실행 상태를 확인하세요');
      }
    } catch (e) {
      toast.error('출력 실패: ' + String(e));
    } finally {
      setBusy(false);
    }
  }

  async function handleDownload() {
    downloadZpl(await renderRepairTagsZpl(items), `tag-${asId || 'repair'}.zpl`);
  }

  const step = (d: number) => setCount((c) => Math.max(1, Math.min(99, c + d)));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" {...backdropClose(onClose)}>
      <EscClose onClose={onClose} />
      <div className="bg-white rounded-xl shadow-2xl w-[440px] max-w-[95vw]" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">꼬리표 출력 · {name} 님</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="p-5 space-y-4">
          {/* 수량 — 실제 도착한 가위 수와 맞는지 확인 후 출력 */}
          <div className="flex items-center justify-between">
            <div>
              <p className="text-xs font-semibold text-neutral-700">가위 수량</p>
              <p className="text-[11px] text-neutral-400">도착한 실물 수와 같은지 확인하세요</p>
            </div>
            <div className="flex items-center gap-2">
              <button onClick={() => step(-1)} className="w-9 h-9 rounded-lg border border-neutral-200 flex items-center justify-center hover:bg-neutral-50"><Minus size={14} /></button>
              <span className="w-10 text-center text-lg font-bold tabular-nums">{count}</span>
              <button onClick={() => step(1)} className="w-9 h-9 rounded-lg border border-neutral-200 flex items-center justify-center hover:bg-neutral-50"><Plus size={14} /></button>
            </div>
          </div>

          {/* 미리보기 — 한 칸에 2개(오른쪽은 라벨지 배치대로 뒤집혀 출력) */}
          <div>
            <p className="text-[10px] text-neutral-400 mb-1">미리보기 · 한 칸에 꼬리표 2개 (오른쪽은 뒤집혀 나옴 — 정상)</p>
            <div className="rounded-lg border border-neutral-200 bg-neutral-100 p-2 space-y-1 max-h-64 overflow-y-auto">
              {previews.length === 0 && <p className="text-xs text-neutral-400 py-6 text-center">렌더링 중...</p>}
              {previews.map((src, i) => (
                // eslint-disable-next-line @next/next/no-img-element
                <img key={i} src={src} alt={`꼬리표 ${i + 1}칸`} className="w-full bg-white" style={{ imageRendering: 'pixelated' }} />
              ))}
            </div>
          </div>

          <p className="text-[11px] text-neutral-500 leading-relaxed">
            유포 꼬리표 라벨지 + 리본이 끼워져 있어야 합니다. 롤을 바꿨다면 헤드를 닫을 때 프린터가 간격을 자동으로 잽니다.
          </p>
        </div>

        <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-neutral-200">
          <Button variant="ghost" size="sm" onClick={handleDownload}><FileDown size={14} />ZPL 다운로드</Button>
          <Button size="sm" onClick={handlePrint} loading={busy}><Printer size={14} />꼬리표 {count}장 출력</Button>
        </div>
      </div>
    </div>
  );
}
