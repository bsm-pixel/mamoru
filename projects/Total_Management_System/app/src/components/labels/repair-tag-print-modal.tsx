'use client';

import { useEffect, useMemo, useState } from 'react';
import { Button } from '@/components/ui/button';
import { X, Printer, FileDown, Minus, Plus } from 'lucide-react';
import toast from 'react-hot-toast';
import { buildRepairTagItems, renderRepairTagsPreview, renderRepairTagsZpl, type RepairTagItem } from '@/lib/label/repair-tag';
import { sendZplToPrinter, downloadZpl } from '@/lib/label/browser-print';
import { backdropClose } from '@/lib/ui/backdrop';
import { EscClose } from '@/components/ui/esc-close';

export interface RepairTagEntry {
  id: string;
  name: string;
  asId: string;
  /** 기본 수량 — 수리내역서가 있으면 그 가위 수, 없으면 접수 수량(마모루+타사) */
  defaultCount: number;
}

interface Props {
  /** 1건(상세 화면) 또는 여러 건(목록에서 선택) — 여러 건이면 빈칸 없이 이어서 출력 */
  entries: RepairTagEntry[];
  onClose: () => void;
}

/** 복원수리 꼬리표 출력 — 건별 수량 확인 후 한 번에 출력 (유포 꼬리표 라벨지 + 리본) */
export function RepairTagPrintModal({ entries, onClose }: Props) {
  const [counts, setCounts] = useState<Record<string, number>>(() =>
    Object.fromEntries(entries.map((e) => [e.id, Math.max(1, e.defaultCount || 1)])),
  );
  const [previews, setPreviews] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const multi = entries.length > 1;

  // 건 순서대로 이어 붙임 → 2개씩 한 칸이라 건 사이에 빈칸이 생기지 않는다(라벨지 절약)
  const items = useMemo<RepairTagItem[]>(
    () => entries.flatMap((e) => buildRepairTagItems(e.name, e.asId, counts[e.id] ?? 1)),
    [entries, counts],
  );
  const total = items.length;
  const units = Math.ceil(total / 2);

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
        toast.success(`꼬리표 ${total}장 출력 전송됨`);
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
    downloadZpl(await renderRepairTagsZpl(items), multi ? `tags-${entries.length}건.zpl` : `tag-${entries[0]?.asId || 'repair'}.zpl`);
  }

  const step = (id: string, d: number) => setCounts((c) => ({ ...c, [id]: Math.max(1, Math.min(99, (c[id] ?? 1) + d)) }));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" {...backdropClose(onClose)}>
      <EscClose onClose={onClose} />
      <div className="bg-white rounded-xl shadow-2xl w-[460px] max-w-[95vw] max-h-[90vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">
            꼬리표 출력 · {multi ? `${entries.length}건` : `${entries[0]?.name ?? ''} 님`}
          </h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="p-5 space-y-4 overflow-y-auto">
          {/* 건별 수량 — 도착한 실물 수와 맞는지 확인 후 출력 */}
          <div>
            <p className="text-xs font-semibold text-neutral-700">가위 수량</p>
            <p className="text-[11px] text-neutral-400 mb-2">도착한 실물 수와 같은지 확인하세요</p>
            <div className="divide-y divide-neutral-100 border-y border-neutral-100">
              {entries.map((e) => (
                <div key={e.id} className="flex items-center justify-between gap-3 py-2">
                  <div className="min-w-0">
                    <p className="text-sm font-semibold text-neutral-800 truncate">{e.name} 님</p>
                    <p className="text-[11px] text-neutral-400">{e.asId}</p>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <button onClick={() => step(e.id, -1)} aria-label={`${e.name} 수량 줄이기`} className="w-8 h-8 rounded-lg border border-neutral-200 flex items-center justify-center hover:bg-neutral-50"><Minus size={13} /></button>
                    <span className="w-8 text-center text-base font-bold tabular-nums">{counts[e.id] ?? 1}</span>
                    <button onClick={() => step(e.id, 1)} aria-label={`${e.name} 수량 늘리기`} className="w-8 h-8 rounded-lg border border-neutral-200 flex items-center justify-center hover:bg-neutral-50"><Plus size={13} /></button>
                  </div>
                </div>
              ))}
            </div>
            <p className="text-[11px] text-neutral-500 mt-2 tabular-nums">
              꼬리표 {total}장 · 라벨 {units}칸{total % 2 === 1 ? ' (마지막 칸 오른쪽 1개는 비어서 나옴)' : ''}
            </p>
          </div>

          {/* 미리보기 — 한 칸에 2개(오른쪽은 라벨지 배치대로 뒤집혀 출력) */}
          <div>
            <p className="text-[10px] text-neutral-400 mb-1">미리보기 · 한 칸에 꼬리표 2개, 건 사이 빈칸 없이 이어서 (오른쪽은 뒤집혀 나옴 — 정상)</p>
            <div className="rounded-lg border border-neutral-200 bg-neutral-100 p-2 space-y-1 max-h-64 overflow-y-auto">
              {previews.length === 0 && <p className="text-xs text-neutral-400 py-6 text-center">렌더링 중...</p>}
              {previews.map((src, i) => (
                // eslint-disable-next-line @next/next/no-img-element
                <img key={i} src={src} alt={`꼬리표 ${i + 1}칸`} className="w-full bg-white" style={{ imageRendering: 'pixelated' }} />
              ))}
            </div>
          </div>

          <p className="text-[11px] text-neutral-500 leading-relaxed">
            유포 꼬리표 라벨지 + 리본이 끼워져 있어야 합니다. 출력 후 FEED 한 번 → 빈 칸에서 뜯기.
          </p>
        </div>

        <div className="flex items-center justify-end gap-2 px-5 py-3 border-t border-neutral-200">
          <Button variant="ghost" size="sm" onClick={handleDownload}><FileDown size={14} />ZPL 다운로드</Button>
          <Button size="sm" onClick={handlePrint} loading={busy}><Printer size={14} />꼬리표 {total}장 출력</Button>
        </div>
      </div>
    </div>
  );
}
