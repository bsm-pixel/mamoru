'use client';

/**
 * 시각적 라벨 디자이너 — 요소 드래그 배치 + 정밀조정 + 실시간 미리보기 + 저장.
 * 크기별 템플릿 자유 생성(빗 등 다른 사이즈) + 변수({product}/{sku}/{serial}) 자동.
 * 저장본 = settings(label.templates). 탭은 templateId key remount로 desync 방지.
 */

import { useEffect, useRef, useState } from 'react';
import { Button } from '@/components/ui/button';
import { ensureLabelFonts, renderLabelPreview, FONT_OPTIONS } from '@/lib/label/render-label';
import { LABEL_TEMPLATES, type LabelTemplate, type LabelElement, type LabelData } from '@/lib/label/templates';
import { useLabelTemplate, useLabelTemplates, useSaveLabelTemplate } from '@/hooks/use-label-templates';
import { Save, RotateCcw, Move, Plus, Trash2 } from 'lucide-react';

const SCALE = 16; // px/mm 편집 배율
const LOGO_ASPECT = 8.8;
const VARS = ['{product}', '{sku}', '{serial}'];

const SAMPLE_BASE: LabelData = { product: 'A2-58DRY-1', sku: 'IW-91', serial: 'MR2610816' };

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
function EditorInner({ templateId, onDeleted }: { templateId: string; onDeleted: () => void }) {
  const baseTpl = useLabelTemplate(templateId);
  const { save, reset, isPending } = useSaveLabelTemplate();
  const isBuiltin = templateId in LABEL_TEMPLATES;

  const [tpl, setTpl] = useState<LabelTemplate>(baseTpl);
  const [dirty, setDirty] = useState(false);
  const [sel, setSel] = useState(-1);
  const [preview, setPreview] = useState('');
  const canvasRef = useRef<HTMLDivElement>(null);
  const dimsRef = useRef({ w: baseTpl.widthMm, h: baseTpl.heightMm });
  dimsRef.current = { w: tpl.widthMm, h: tpl.heightMm };

  const baseJson = JSON.stringify(baseTpl);
  const appliedRef = useRef<string>(baseJson);
  useEffect(() => {
    if (!dirty && baseJson !== appliedRef.current) { appliedRef.current = baseJson; setTpl(JSON.parse(baseJson) as LabelTemplate); }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [baseJson, dirty]);

  useEffect(() => {
    let alive = true;
    const t = setTimeout(async () => {
      await ensureLabelFonts();
      if (!alive) return;
      try { setPreview(await renderLabelPreview(tpl, SAMPLE_BASE, 300)); } catch { /* ignore */ }
    }, 90);
    return () => { alive = false; clearTimeout(t); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tpl]);

  const W = tpl.widthMm * SCALE, H = tpl.heightMm * SCALE;

  function patchEl(i: number, patch: Partial<LabelElement>) {
    setDirty(true);
    setTpl((t) => ({ ...t, elements: t.elements.map((e, j) => (j === i ? { ...e, ...patch } : e)) }));
  }
  function patchTpl(patch: Partial<LabelTemplate>) { setDirty(true); setTpl((t) => ({ ...t, ...patch })); }
  function addEl(el: LabelElement) { setDirty(true); setSel(tpl.elements.length); setTpl((t) => ({ ...t, elements: [...t.elements, el] })); }
  function deleteEl(i: number) { setDirty(true); setSel(-1); setTpl((t) => ({ ...t, elements: t.elements.filter((_, j) => j !== i) })); }

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
  function handleResetOrDelete() {
    reset(templateId);
    if (isBuiltin) { setTpl(LABEL_TEMPLATES[templateId]); setDirty(false); setSel(-1); }
    else { onDeleted(); }
  }

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
            <Button variant="ghost" size="sm" onClick={handleResetOrDelete} className={isBuiltin ? '' : 'text-red-500'}>
              {isBuiltin ? <><RotateCcw size={13} />기본값</> : <><Trash2 size={13} />템플릿 삭제</>}
            </Button>
            <Button size="sm" onClick={handleSave} loading={isPending} disabled={!dirty && !isPending}><Save size={13} />저장</Button>
          </div>
        </div>

        {/* 라벨 크기 */}
        <div className="flex items-center gap-3 rounded-lg bg-neutral-50 border border-neutral-200 px-3 py-2 text-xs">
          <span className="font-semibold text-neutral-500">라벨 크기</span>
          <span>가로</span>
          <input type="number" value={tpl.widthMm} onChange={(e) => patchTpl({ widthMm: Math.max(10, parseFloat(e.target.value) || 10) })} className="w-14 h-7 px-1.5 text-center border border-neutral-200 rounded" />
          <span>×</span>
          <span>세로</span>
          <input type="number" value={tpl.heightMm} onChange={(e) => patchTpl({ heightMm: Math.max(8, parseFloat(e.target.value) || 8) })} className="w-14 h-7 px-1.5 text-center border border-neutral-200 rounded" />
          <span className="text-neutral-400">mm</span>
        </div>

        {/* 요소 칩 */}
        <div className="flex flex-wrap gap-1">
          {tpl.elements.map((el, i) => (
            <button key={i} onClick={() => setSel(i)}
              className={`px-2.5 py-1 rounded-md text-xs transition ${sel === i ? 'bg-blue-600 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'}`}>
              {elemLabel(el)}
            </button>
          ))}
        </div>

        {/* 요소 추가 */}
        <div className="flex items-center gap-2 rounded-lg bg-neutral-50 border border-neutral-200 px-3 py-2">
          <span className="text-[11px] font-semibold text-neutral-500">요소 추가</span>
          <button onClick={() => addEl({ kind: 'text', xMm: 3, yMm: tpl.heightMm * 0.45, text: '새 텍스트', fontFamily: 'Inter', weight: 700, sizeMm: 2.5 })}
            className="px-3 py-1 rounded-md text-xs font-medium bg-white border border-neutral-300 text-neutral-700 hover:bg-neutral-100 flex items-center gap-1"><Plus size={12} /> 텍스트</button>
          <button onClick={() => addEl({ kind: 'rule', xMm: 2, yMm: tpl.heightMm * 0.5, lengthMm: tpl.widthMm * 0.5, thicknessMm: 0.4 })}
            className="px-3 py-1 rounded-md text-xs font-medium bg-white border border-neutral-300 text-neutral-700 hover:bg-neutral-100 flex items-center gap-1"><Plus size={12} /> 구분선</button>
          <button onClick={() => addEl({ kind: 'barcode', xMm: 2, yMm: tpl.heightMm * 0.55, data: '{sku}', widthMm: Math.min(26, tpl.widthMm * 0.7), heightMm: 6, showText: false })}
            className="px-3 py-1 rounded-md text-xs font-medium bg-white border border-neutral-300 text-neutral-700 hover:bg-neutral-100 flex items-center gap-1"><Plus size={12} /> 바코드</button>
        </div>

        {selEl ? (
          <div className="space-y-2.5 rounded-xl border border-neutral-200 p-4">
            <div className="flex items-center justify-between">
              <p className="text-xs font-semibold text-neutral-700">{elemLabel(selEl)}</p>
              <button onClick={() => deleteEl(sel)} className="text-[11px] text-red-500 hover:text-red-700 flex items-center gap-0.5"><Trash2 size={11} />삭제</button>
            </div>
            {selEl.kind === 'text' && (
              <>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs text-neutral-500 w-16">내용</span>
                  <input type="text" value={selEl.text || ''} onChange={(e) => patchEl(sel, { text: e.target.value })}
                    className="flex-1 h-7 px-2 text-xs border border-neutral-200 rounded" placeholder="고정 글씨 또는 변수" />
                </div>
                <div className="flex items-center gap-1 flex-wrap pl-16">
                  <span className="text-[10px] text-neutral-400">변수 삽입:</span>
                  {VARS.map((v) => (
                    <button key={v} onClick={() => patchEl(sel, { text: (selEl.text || '') + v })}
                      className="px-1.5 py-0.5 rounded text-[10px] font-mono bg-blue-50 text-blue-600 hover:bg-blue-100">{v}</button>
                  ))}
                </div>
              </>
            )}
            {selEl.kind === 'barcode' && (
              <div className="flex items-center justify-between gap-2">
                <span className="text-xs text-neutral-500 w-16">바코드값</span>
                <input type="text" value={selEl.data || ''} onChange={(e) => patchEl(sel, { data: e.target.value })}
                  className="flex-1 h-7 px-2 text-xs border border-neutral-200 rounded font-mono" placeholder="{sku} 또는 {serial}" />
              </div>
            )}
            <Field label="가로(X)" value={selEl.xMm} onChange={(v) => patchEl(sel, { xMm: v })} />
            <Field label="세로(Y)" value={selEl.yMm} onChange={(v) => patchEl(sel, { yMm: v })} />
            {selEl.kind === 'text' && (
              <>
                <Field label="글자크기" value={selEl.sizeMm || 2} onChange={(v) => patchEl(sel, { sizeMm: Math.max(0.5, v) })} step={0.2} />
                <div className="flex items-center justify-between gap-2">
                  <span className="text-xs text-neutral-500 w-16">폰트</span>
                  <select value={selEl.fontFamily || 'Inter'} onChange={(e) => patchEl(sel, { fontFamily: e.target.value })}
                    className="h-7 px-2 text-xs border border-neutral-200 rounded flex-1">
                    {FONT_OPTIONS.map((f) => <option key={f} value={f}>{f}</option>)}
                  </select>
                </div>
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
  const all = useLabelTemplates();
  const { save } = useSaveLabelTemplate();
  const ids = Object.keys(all);
  const [templateId, setTemplateId] = useState(ids[0] || 'product_40x20');
  const [creating, setCreating] = useState(false);
  const [newName, setNewName] = useState('');
  const [newW, setNewW] = useState(40);
  const [newH, setNewH] = useState(30);

  function createTemplate() {
    const id = `custom_${Date.now()}`;
    const tpl: LabelTemplate = {
      id,
      name: newName.trim() || `새 라벨 ${newW}×${newH}`,
      widthMm: newW,
      heightMm: newH,
      elements: [
        { kind: 'image', xMm: 2, yMm: 2, src: '/labels/mamoru-logo.png', widthMm: Math.min(13, newW * 0.5) },
        { kind: 'text', xMm: 2, yMm: Math.round(newH * 0.4), text: '{product}', fontFamily: 'Inter', weight: 700, sizeMm: Math.max(2, Math.round(newH * 0.16)) },
        { kind: 'barcode', xMm: 2, yMm: Math.round(newH * 0.65), data: '{sku}', widthMm: Math.min(26, newW * 0.7), heightMm: Math.max(4, Math.round(newH * 0.28)), showText: false },
      ],
    };
    save(tpl);
    setTemplateId(id);
    setCreating(false);
    setNewName('');
  }

  return (
    <div className="space-y-3">
      <div className="flex gap-1 flex-wrap items-center">
        {Object.values(all).map((t) => (
          <button key={t.id} onClick={() => setTemplateId(t.id)}
            className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition ${templateId === t.id ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
            {t.name}
          </button>
        ))}
        <button onClick={() => setCreating((v) => !v)}
          className="px-3 py-1.5 rounded-lg text-xs font-semibold border border-dashed border-neutral-300 text-neutral-500 hover:bg-neutral-50 flex items-center gap-1"><Plus size={12} /> 새 템플릿</button>
      </div>

      {creating && (
        <div className="flex items-center gap-2 rounded-lg bg-blue-50 border border-blue-200 px-3 py-2 text-xs flex-wrap">
          <input type="text" value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="템플릿 이름 (예: 빗 라벨)" className="h-7 px-2 border border-neutral-200 rounded w-40" />
          <span>가로</span><input type="number" value={newW} onChange={(e) => setNewW(Math.max(10, parseInt(e.target.value) || 10))} className="w-14 h-7 px-1.5 text-center border border-neutral-200 rounded" />
          <span>×</span>
          <span>세로</span><input type="number" value={newH} onChange={(e) => setNewH(Math.max(8, parseInt(e.target.value) || 8))} className="w-14 h-7 px-1.5 text-center border border-neutral-200 rounded" />
          <span className="text-neutral-400">mm</span>
          <Button size="sm" onClick={createTemplate}>만들기</Button>
          <Button size="sm" variant="ghost" onClick={() => setCreating(false)}>취소</Button>
        </div>
      )}

      <EditorInner key={templateId} templateId={templateId} onDeleted={() => setTemplateId('product_40x20')} />
    </div>
  );
}
