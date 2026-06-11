'use client';

import { useEffect, useState } from 'react';
import { Button } from '@/components/ui/button';
import { X, Printer, FileDown } from 'lucide-react';
import toast from 'react-hot-toast';
import type { LabelTemplate, LabelData } from '@/lib/label/templates';
import { ensureLabelFonts, renderLabelPreview, renderLabelZpl } from '@/lib/label/render-label';
import { sendZplToPrinter, downloadZpl } from '@/lib/label/browser-print';

interface Props {
  template: LabelTemplate;
  data: LabelData;
  title?: string;
  onClose: () => void;
}

const DPI = 203; // ZT231

export function LabelPrintModal({ template, data, title, onClose }: Props) {
  const [copies, setCopies] = useState(1);
  const [preview, setPreview] = useState<string>('');
  const [busy, setBusy] = useState(false);

  // 폰트 로드 후 미리보기 렌더 (고배율로 떠서 선명)
  useEffect(() => {
    let alive = true;
    (async () => {
      await ensureLabelFonts();
      if (!alive) return;
      try { setPreview(renderLabelPreview(template, data, 300)); } catch { /* ignore */ }
    })();
    return () => { alive = false; };
  }, [template, data]);

  async function handlePrint() {
    setBusy(true);
    try {
      await ensureLabelFonts();
      const zpl = renderLabelZpl(template, data, DPI, copies);
      const ok = await sendZplToPrinter(zpl);
      if (ok) {
        toast.success(`라벨 ${copies}장 출력 전송됨`);
        onClose();
      } else {
        toast.error('프린터(Browser Print) 미연결 — ZPL 다운로드로 테스트하세요');
      }
    } catch (e) {
      toast.error('출력 실패: ' + String(e));
    } finally {
      setBusy(false);
    }
  }

  async function handleDownload() {
    await ensureLabelFonts();
    const zpl = renderLabelZpl(template, data, DPI, copies);
    downloadZpl(zpl, `${template.id}-${(data.serial || data.sku || 'label')}.zpl`);
    toast.success('ZPL 다운로드 — ZSU 등으로 프린터에 보내 테스트하세요');
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl w-[420px] max-w-[95vw]" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">{title || '라벨 출력'}</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        <div className="p-5 space-y-4">
          {/* 미리보기 — 40×20 비율, 실제 출력과 동일(WYSIWYG) */}
          <div>
            <p className="text-[10px] text-neutral-400 mb-1">미리보기 ({template.widthMm}×{template.heightMm}mm · 실제 출력과 동일)</p>
            <div className="rounded-lg border border-neutral-200 bg-neutral-100 p-3 flex items-center justify-center" style={{ minHeight: 120 }}>
              {preview ? (
                // eslint-disable-next-line @next/next/no-img-element
                <img src={preview} alt="label preview" style={{ width: '100%', maxWidth: 340, imageRendering: 'pixelated', background: '#fff' }} />
              ) : (
                <span className="text-xs text-neutral-400">렌더링 중...</span>
              )}
            </div>
          </div>

          {/* 매수 */}
          <div className="flex items-center gap-2">
            <label className="text-xs text-neutral-500">매수</label>
            <input type="number" min={1} max={200} value={copies}
              onChange={(e) => setCopies(Math.max(1, Math.min(200, parseInt(e.target.value) || 1)))}
              className="w-20 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm" />
            <span className="text-[11px] text-neutral-400">장</span>
          </div>
        </div>

        <div className="flex justify-end gap-2 px-5 py-3 border-t border-neutral-200">
          <Button variant="ghost" size="sm" onClick={handleDownload}><FileDown size={14} />ZPL 다운로드</Button>
          <Button size="sm" onClick={handlePrint} loading={busy}><Printer size={14} />프린터 출력</Button>
        </div>
      </div>
    </div>
  );
}
