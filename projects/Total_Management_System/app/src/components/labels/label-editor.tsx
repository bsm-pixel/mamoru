'use client';

/**
 * 시각적 라벨 편집기 — 요소를 드래그로 배치 + 정밀(±0.5mm) 조정 + 실시간 미리보기 + 저장.
 * 저장본은 settings(label.templates)에 → 출력 시 그 값 사용. 사장님이 직접 레이아웃 디자인.
 * 탭(템플릿) 전환은 key 기반 remount로 상태 완전 초기화(미리보기·테두리 desync 방지).
 */

import { useEffect, useRef, useState } from 'react';
import { Button } from '@/components/ui/button';
import { ensureLabelFonts, renderLabelPreview } from '@/lib/label/render-label';
import { LABEL_TEMPLATES, type LabelTemplate, type LabelElement, type LabelData } from '@/lib/label/templates';
import { useLabelTemplate, useSaveLabelTemplate } from '@/hooks/use-label-templates';
import { Save, RotateCcw, Move } from 'lucide-react';

const SCALE = 16; // px/mm 편집 배율
const LOGO_ASPECT = 8.8;

const SAMPLE: Record<string, LabelData> = {
  product_40x20: { product: 'A2-58DRY-1', sku: 'IW-91' },
  serial_40x20: { product: 'A2-58DRY-1', serial: 'MR2610816' },
};

function elemLabel(el: LabelElement): string {
  if (el.kind === 'image') return '로고';
  if (el.kind === 'rule') return '구분선';
  if (el.kind === 'barcode') return `바코드(${el.data})`;
  return el.text || '텍스트';
}

function elemBox(el: LabelElement): { w: number; h: number } {
  if (el.kind === 'image') { const w = el.widthMm || 12; return { w, h: el.heightMm || w / LOGO_ASPECT }; }
  if (el.kind === 'barcode') return { w: el.widthMm || 25, h: el.heightMm || 6 };
  if (el.kind === 'rule') return { w: el.lengthMm || 10, h: Math.max(el.thicknessMm || 0.4, 1.2) };
  const s = el.sizeMm || 2;
  return { w: Math.max(4, (el.text || '').length * s * 0.62), h: s * 1.4 };
}

const round1 = (n: number) => Math.round(n * 10) / 10;

function Field({ label, value, onChange, step = 0.5, suffix = 'mm' }: { label: string; value: number; onChange: (v: number) => void; step?: number; suffix?: string }) {
  return (
    <div className="flex items-center justify-between gap-2">
      <span className="text-xs text-neutral-500 w-16">{label}</span>
      <div className="flex items-center gap-1">
        <button onClick={() => onChange(round1(value - step))} className="w-6 h-6 rounded border border-neutral-200 text-neutral-600 hover:bg-neutral-100">−</button>
        <input type="number" value={value} step={step} onChange={(e) => onChange(round1(parseFloat(e.target.value) || 0))}
          className="w-16 h-7 px-2 text-center text-xs border border-neutral-200 rounded" />
        <button onClick={() => onChange(round1(value + step))} className="w-6 h-6 rounded border border-neutral-200 text-neutral-600 hover:bg-neutral-100">+</button>
        <span className="text-[10px] text-neutral-400 w-5">{suffix}</span>
      </div>
    </div>
  );
}

/** 단일 템플릿 편집 영역 (templateId별 key remount) */
function EditorInner({ templateId }: { templateId: string }) {
  const baseTpl = useLabelTemplate(templateId);
  const { save, reset, isPending } = useSaveLabelTemplate();

  const [tpl, setTpl] = useState<LabelTemplate>(baseTpl);
  const [dirty, setDirty] = useState(false);
  const [sel, setSel] = useState(-1);
  const [preview, setPreview] = useState('');
  const canvasRef = useRef<HTMLDivElement>(null);
  const dimsRef = useRef({ w: baseTpl.widthMm, h: baseTpl.heightMm });
  dimsRef.current = { w: tpl.widthMm, h: tpl.heightMm };

  // 저장본(비동기 settings)이 도착하면 편집 전 한정 반영.
  // ⚠️ useSetting이 매 렌더 새 객체를 반환하므로 identity 의존 시 무한루프 → 직렬화(내용) 기준으로 비교.
  const baseJson = JSON.stringify(baseTpl);
  const appliedRef = useRef<string>(baseJson);
  useEffect(() => {
    if (!dirty && baseJson !== appliedRef.current) {
      appliedRef.current = baseJson;
      setTpl(JSON.parse(baseJson) as LabelTemplate);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [baseJson, dirty]);

  const sample = SAMPLE[templateId] || {};

  // 미리보기 재렌더(디바운스) — tpl 따라감
  useEffect(() => {
    let alive = true;
    const t = setTimeout(async () => {
      await ensureLabelFonts();
      if (!alive) return;
      try { setPreview(await renderLabelPreview(tpl, sample, 300)); } catch { /* ignore */ }
    }, 90);
    return () => { alive = false; clearTimeout(t); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tpl]);

  const W = tpl.widthMm * SCALE, H = tpl.heightMm * SCALE;

  function patchEl(i: number, patch: Partial<LabelElement>) {
    setDirty(true);
    setTpl((t) => ({ ...t, elements: t.elements.map((e, j) => (j === i ? { ...e, ...patch } : e)) }));
  }

  function startDrag(e: React.MouseEvent, i: number) {
    e.preventDefault();
    setSel(i);
    const rect = canvasRef.current!.getBoundingClientRect();
    const el = tpl.elements[i];
    const offX = (e.clientX - rect.left) / SCALE - el.xMm;
    const offY = (e.clientY - rect.top) / SCALE - el.yMm;
    const move = (ev: MouseEvent) => {
      const r = canvasRef.current!.getBoundingClientRect();
      const nx = round1(Math.max(0, Math.min(dimsRef.current.w, (ev.clientX - r.left) / SCALE - offX)));
      const ny = round1(Math.max(0, Math.min(dimsRef.current.h, (ev.clientY - r.top) / SCALE - offY)));
      setDirty(true);
      setTpl((t) => ({ ...t, elements: t.elements.map((e2, j) => (j === i ? { ...e2, xMm: nx, yMm: ny } : e2)) }));
    };
    const up = () => { window.removeEventListener('mousemove', move); window.removeEventListener('mouseup', up); };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  }

  function handleSave() { save(tpl); setDirty(false); }
  function handleReset() { reset(templateId); setTpl(LABEL_TEMPLATES[templateId]); setDirty(false); setSel(-1); }

  const selEl = sel >= 0 ? tpl.elements[sel] : null;

  return (
    <div className="flex flex-col lg:flex-row gap-6">
      {/* 캔버스 */}
      <div className="space-y-3">
        <div className="inline-block rounded-lg border border-neutral-300 bg-white p-2 shadow-sm">
          <div ref={canvasRef} className="relative" style={{ width: W, height: H, background: '#fff', outline: '1px dashed #d4d4d4' }}>
            {preview && (
              // eslint-disable-next-line @next/next/no-img-element
              <img src={preview} alt="" className="absolute inset-0" style={{ width: W, height: H, imageRendering: 'pixelated' }} draggable={false} />
            )}
            {tpl.elements.map((el, i) => {
              const b = elemBox(el);
              return (
                <div key={i} onMouseDown={(e) => startDrag(e, i)} onClick={() => setSel(i)}
                  className="absolute cursor-move"
                  style={{
                    left: el.xMm * SCALE, top: el.yMm * SCALE, width: b.w * SCALE, height: b.h * SCALE,
                    border: sel === i ? '1.5px solid #2563eb' : '1px dashed rgba(120,120,120,0.45)',
                    background: sel === i ? 'rgba(37,99,235,0.08)' : 'transparent',
                  }}
                  title={elemLabel(el)}
                />
              );
            })}
          </div>
        </div>
        <p className="text-[11px] text-neutral-400 flex items-center gap-1"><Move size={12} /> 요소를 드래그해 배치 · 클릭 후 오른쪽에서 정밀 조정 ({tpl.widthMm}×{tpl.heightMm}mm)</p>
      </div>

      {/* 속성 패널 */}
      <div className="flex-1 min-w-[260px] space-y-4">
        <div className="flex items-center justify-between">
          <h3 className="text-sm font-bold text-neutral-800">요소 조정</h3>
          <div className="flex gap-2">
            <Button variant="ghost" size="sm" onClick={handleReset}><RotateCcw size={13} />기본값</Button>
            <Button size="sm" onClick={handleSave} loading={isPending} disabled={!dirty && !isPending}><Save size={13} />저장</Button>
          </div>
        </div>

        <div className="flex flex-wrap gap-1">
          {tpl.elements.map((el, i) => (
            <button key={i} onClick={() => setSel(i)}
              className={`px-2.5 py-1 rounded-md text-xs transition ${sel === i ? 'bg-blue-600 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'}`}>
              {elemLabel(el)}
            </button>
          ))}
        </div>

        {selEl ? (
          <div className="space-y-2.5 rounded-xl border border-neutral-200 p-4">
            <p className="text-xs font-semibold text-neutral-700">{elemLabel(selEl)}</p>
            <Field label="가로(X)" value={selEl.xMm} onChange={(v) => patchEl(sel, { xMm: v })} />
            <Field label="세로(Y)" value={selEl.yMm} onChange={(v) => patchEl(sel, { yMm: v })} />
            {selEl.kind === 'text' && (
              <>
                <Field label="글자크기" value={selEl.sizeMm || 2} onChange={(v) => patchEl(sel, { sizeMm: Math.max(0.5, v) })} step={0.2} />
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs text-neutral-500 w-16">굵기</span>
                  <select value={selEl.weight || 400} onChange={(e) => patchEl(sel, { weight: parseInt(e.target.value) })}
                    className="h-7 px-2 text-xs border border-neutral-200 rounded flex-1">
                    {[400, 500, 600, 700, 800, 900].map((w) => <option key={w} value={w}>{w}</option>)}
                  </select>
                </div>
              </>
            )}
            {selEl.kind === 'image' && (
              <Field label="크기(폭)" value={selEl.widthMm || 12} onChange={(v) => patchEl(sel, { widthMm: Math.max(2, v) })} />
            )}
            {selEl.kind === 'barcode' && (
              <>
                <Field label="바코드폭" value={selEl.widthMm || 25} onChange={(v) => patchEl(sel, { widthMm: Math.max(5, v) })} />
                <Field label="바코드높이" value={selEl.heightMm || 6} onChange={(v) => patchEl(sel, { heightMm: Math.max(2, v) })} />
              </>
            )}
            {selEl.kind === 'rule' && (
              <>
                <Field label="선길이" value={selEl.lengthMm || 10} onChange={(v) => patchEl(sel, { lengthMm: Math.max(1, v) })} />
                <Field label="선굵기" value={selEl.thicknessMm || 0.4} onChange={(v) => patchEl(sel, { thicknessMm: Math.max(0.1, v) })} step={0.1} />
              </>
            )}
          </div>
        ) : (
          <p className="text-xs text-neutral-400 rounded-xl border border-dashed border-neutral-200 p-4 text-center">캔버스에서 요소를 클릭하세요</p>
        )}

        {dirty && <p className="text-[11px] text-amber-600">● 변경됨 — 저장을 눌러야 출력에 반영됩니다</p>}
      </div>
    </div>
  );
}

export function LabelEditor() {
  const [templateId, setTemplateId] = useState('product_40x20');
  return (
    <div className="space-y-3">
      <div className="flex gap-1">
        {Object.values(LABEL_TEMPLATES).map((t) => (
          <button key={t.id} onClick={() => setTemplateId(t.id)}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition ${templateId === t.id ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
            {t.name}
          </button>
        ))}
      </div>
      <EditorInner key={templateId} templateId={templateId} />
    </div>
  );
}
