'use client';

import { useState, useEffect } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useSaveInspections } from '@/hooks/use-repairs';
import type { RepairInspection } from '@/lib/supabase/types';
import { Plus, Trash2, Save, Camera, X, Loader2 } from 'lucide-react';
import toast from 'react-hot-toast';
import { resizeImage } from '@/lib/utils/resize-image';
import { createClient } from '@/lib/supabase/client';
import { InspectionMarkBoard } from './inspection-mark-board';
import { type MarkV2 } from './inspection-marks';
import { COMMENT_PRESETS_BY_TYPE } from '@/lib/repair/comment-presets';
import { applyMarkMents, applyFlagMents, splitBlocks } from './ment-linkage';

const SCISSOR_TYPES = ['블런트', '틴닝', '장가위', '슬라이싱', '기타'];

interface InspectionItem {
  scissor_number: number;
  scissor_type: string;
  photo_url: string;
  marks: MarkV2[];
  flags: string[];
  comment: string;
  worker: string;
  _photoPreview?: string;
}

function newInspectionItem(num: number): InspectionItem {
  return { scissor_number: num, scissor_type: '블런트', photo_url: '', marks: [], flags: [], comment: '', worker: '백성민' };
}

/** repair_inspections.photo_marks(점·선·플래그 혼합) → marks/flags 분리 */
function splitMarks(raw: unknown): { marks: MarkV2[]; flags: string[] } {
  const arr = Array.isArray(raw) ? (raw as Array<{ label: string; x?: number; y?: number; x2?: number; y2?: number; flag?: boolean }>) : [];
  const marks = arr.filter((m) => !m.flag && typeof m.x === 'number').map((m) => ({ label: m.label, x: m.x as number, y: m.y as number, x2: m.x2, y2: m.y2 }));
  const flags = arr.filter((m) => m.flag).map((m) => m.label);
  return { marks, flags };
}

interface InspectionFormProps {
  repairId: string;
  existingInspections: RepairInspection[];
  totalScissors: number;
  onSaved?: () => void;
}

/** 모바일 가로 감지 (검수 화면 세로 고정) */
function useLandscapeMobile() {
  const [landscape, setLandscape] = useState(false);
  useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(orientation: landscape) and (max-height: 500px)');
    const sync = () => setLandscape(mq.matches);
    sync();
    mq.addEventListener('change', sync);
    return () => mq.removeEventListener('change', sync);
  }, []);
  return landscape;
}

const draftKey = (repairId: string) => `repair_inspect_${repairId}`;
interface Draft { items: InspectionItem[]; activeIdx: number; }
function loadDraft(repairId: string): Draft | null {
  if (typeof window === 'undefined') return null;
  try { const raw = sessionStorage.getItem(draftKey(repairId)); return raw ? (JSON.parse(raw) as Draft) : null; } catch { return null; }
}

export function InspectionForm({ repairId, existingInspections, onSaved }: InspectionFormProps) {
  const draft = typeof window !== 'undefined' ? loadDraft(repairId) : null;

  const [items, setItems] = useState<InspectionItem[]>(() => {
    if (draft?.items?.length) return draft.items;
    if (existingInspections.length > 0) {
      return existingInspections.map((e) => {
        const { marks, flags } = splitMarks(e.photo_marks);
        return {
          scissor_number: e.scissor_number,
          scissor_type: e.scissor_type || '블런트',
          photo_url: e.photo_url || '',
          marks, flags,
          comment: (e as { comment?: string | null }).comment || '',
          worker: e.worker,
        };
      });
    }
    return [newInspectionItem(1)];
  });
  const [activeIdx, setActiveIdx] = useState(draft?.activeIdx ?? 0);
  const [uploading, setUploading] = useState(false);

  const saveInspections = useSaveInspections();
  const landscapeMobile = useLandscapeMobile();

  // 작성 중 draft 보존 (카메라/포커스 손실로 모달 재마운트 시 복원)
  useEffect(() => {
    if (typeof window === 'undefined') return;
    try { sessionStorage.setItem(draftKey(repairId), JSON.stringify({ items, activeIdx })); } catch { /* noop */ }
  }, [repairId, items, activeIdx]);

  const patchActive = (patch: Partial<InspectionItem>) =>
    setItems((prev) => prev.map((it, i) => (i === activeIdx ? { ...it, ...patch } : it)));

  const addScissor = () => { setItems((prev) => [...prev, newInspectionItem(prev.length + 1)]); setActiveIdx(items.length); };
  const removeScissor = (idx: number) => {
    if (items.length <= 1) return;
    setItems((prev) => prev.filter((_, i) => i !== idx).map((it, i) => ({ ...it, scissor_number: i + 1 })));
    setActiveIdx(Math.max(0, activeIdx - 1));
  };

  // 핀/플래그 변경 → 자동 멘트 연동 (해당 가위)
  const handleMarks = (newMarks: MarkV2[]) =>
    setItems((prev) => prev.map((it, i) => (i === activeIdx ? { ...it, marks: newMarks, comment: applyMarkMents(it.comment, it.marks, newMarks) } : it)));
  const handleFlags = (newFlags: string[]) =>
    setItems((prev) => prev.map((it, i) => (i === activeIdx ? { ...it, flags: newFlags, comment: applyFlagMents(it.comment, it.flags, newFlags, it.scissor_type) } : it)));

  const handlePhotoSelect = async (idx: number, file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, _photoPreview: e.target?.result as string } : it)));
    reader.readAsDataURL(file);

    setUploading(true);
    try {
      const resized = await resizeImage(file, 1200, 0.8);
      const supabase = createClient();
      const ext = file.name.split('.').pop() || 'jpg';
      const filePath = `inspections/${repairId}/${Date.now()}_${idx}.${ext}`;
      const { error: uploadErr } = await supabase.storage.from('repair-photos').upload(filePath, resized, { contentType: 'image/jpeg', upsert: false });
      if (uploadErr) throw uploadErr;
      const { data: urlData } = supabase.storage.from('repair-photos').getPublicUrl(filePath);
      // 새 사진 → 마킹·플래그·진단멘트 초기화
      setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, photo_url: urlData.publicUrl, _photoPreview: undefined, marks: [], flags: [], comment: '' } : it)));
    } catch (err) {
      console.error('검수 사진 업로드 실패:', err);
      setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, _photoPreview: undefined } : it)));
      toast.error('사진 업로드 실패 — 다시 시도해주세요');
    } finally {
      setUploading(false);
    }
  };

  const removePhoto = (idx: number) =>
    setItems((prev) => prev.map((it, i) => (i === idx ? { ...it, photo_url: '', _photoPreview: undefined, marks: [], flags: [], comment: '' } : it)));

  // 가위 종류별 수동 멘트 칩 (공통 멘트는 핀/플래그로 자동 삽입)
  const togglePreset = (p: string) => {
    const t = p.trim();
    setItems((prev) => prev.map((it, i) => {
      if (i !== activeIdx) return it;
      const blocks = splitBlocks(it.comment);
      const k = blocks.indexOf(t);
      if (k >= 0) blocks.splice(k, 1); else blocks.push(t);
      return { ...it, comment: blocks.join('\n\n') };
    }));
  };

  const handleSave = () => {
    const rows = items.map((it) => ({
      scissor_number: it.scissor_number,
      scissor_type: it.scissor_type,
      photo_url: it.photo_url || null,
      photo_marks: [...it.marks, ...it.flags.map((note) => ({ label: note, flag: true }))],
      comment: it.comment.trim() || null,
      worker: it.worker,
    }));
    saveInspections.mutate({ repairId, inspections: rows }, {
      onSuccess: () => { try { sessionStorage.removeItem(draftKey(repairId)); } catch { /* noop */ } onSaved?.(); },
    });
  };

  const current = items[activeIdx];
  if (!current) return null;
  const photoSrc = current._photoPreview || current.photo_url;
  const busy = uploading || saveInspections.isPending;
  const presetList = COMMENT_PRESETS_BY_TYPE[current.scissor_type] || [];
  const presetActive = (p: string) => splitBlocks(current.comment).includes(p.trim());

  return (
    <Card>
      {landscapeMobile && (
        <div className="fixed inset-0 z-[70] bg-white flex flex-col items-center justify-center text-center px-8">
          <span className="text-5xl mb-4">📱</span>
          <p className="text-base font-bold text-neutral-800">세로로 돌려주세요</p>
          <p className="text-sm text-neutral-500 mt-1.5">수리내역서 작성은 세로 화면에서 진행해주세요.</p>
        </div>
      )}

      <CardHeader>
        <CardTitle className="flex items-center justify-between">
          <span>수리내역서 작성</span>
          <Button variant="primary" size="sm" onClick={handleSave} disabled={busy} loading={saveInspections.isPending}>
            <Save size={14} />
            {uploading ? '사진 업로드 중...' : '저장'}
          </Button>
        </CardTitle>
      </CardHeader>

      {/* 가위 탭 */}
      <div className="flex gap-1 mb-4 overflow-x-auto">
        {items.map((item, idx) => (
          <button key={idx} onClick={() => setActiveIdx(idx)}
            className={`shrink-0 px-3 py-1.5 text-xs font-medium rounded-lg transition ${activeIdx === idx ? 'bg-neutral-800 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'}`}>
            #{item.scissor_number} {item.scissor_type}
          </button>
        ))}
        <button onClick={addScissor} className="shrink-0 px-3 py-1.5 text-xs text-neutral-400 hover:text-neutral-600 flex items-center gap-1">
          <Plus size={12} /> 추가
        </button>
      </div>

      <div className="space-y-4">
        {/* 종류 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0">종류</label>
          <select value={current.scissor_type} onChange={(e) => patchActive({ scissor_type: e.target.value })}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm">
            {SCISSOR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          {items.length > 1 && (
            <button onClick={() => removeScissor(activeIdx)} className="text-error/60 hover:text-error"><Trash2 size={14} /></button>
          )}
        </div>

        {/* 사진 촬영 */}
        <div className="flex items-start gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0 pt-2">사진</label>
          {photoSrc ? (
            <div className="relative">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src={photoSrc} alt="검수 사진" className="w-28 h-28 object-cover rounded-lg border border-neutral-200" />
              {uploading && (
                <div className="absolute inset-0 bg-black/40 rounded-lg flex items-center justify-center"><Loader2 size={24} className="text-white animate-spin" /></div>
              )}
              <button onClick={() => removePhoto(activeIdx)} className="absolute -top-2 -right-2 w-7 h-7 bg-error text-white rounded-full flex items-center justify-center shadow-md"><X size={14} /></button>
              <label className="absolute bottom-1 right-1 w-7 h-7 bg-white/80 rounded-full flex items-center justify-center cursor-pointer shadow hover:bg-white transition">
                <Camera size={14} className="text-neutral-600" />
                <input type="file" accept="image/*" capture="environment" className="hidden"
                  onChange={(e) => { const f = e.target.files?.[0]; if (f) handlePhotoSelect(activeIdx, f); e.target.value = ''; }} />
              </label>
            </div>
          ) : (
            <label className="cursor-pointer flex flex-col items-center justify-center w-28 h-28 border-2 border-dashed border-neutral-300 rounded-lg hover:border-neutral-400 transition">
              {uploading ? <Loader2 size={24} className="text-neutral-500 mb-1 animate-spin" /> : <Camera size={24} className="text-neutral-400 mb-1" />}
              <span className="text-[11px] text-neutral-400">{uploading ? '업로드중...' : '촬영/업로드'}</span>
              <input type="file" accept="image/*" capture="environment" className="hidden"
                onChange={(e) => { const f = e.target.files?.[0]; if (f) handlePhotoSelect(activeIdx, f); e.target.value = ''; }} />
            </label>
          )}
        </div>

        {/* 핀 마킹 (자동 멘트 연동) */}
        <div>
          <label className="block text-sm text-neutral-500 mb-2">상처 표시</label>
          <InspectionMarkBoard photoUrl={current.photo_url} marks={current.marks} flags={current.flags} onMarks={handleMarks} onFlags={handleFlags} scissorType={current.scissor_type} />
        </div>

        {/* 작업자 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0">작업자</label>
          <input type="text" value={current.worker} onChange={(e) => patchActive({ worker: e.target.value })}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
        </div>

        {/* 진단 및 내역 — 가위별 (종류별 수동 칩 + 핀/플래그 자동) */}
        <div className="pt-3 border-t border-neutral-100">
          <label className="block text-sm font-medium text-neutral-600 mb-2">진단 및 내역 — #{current.scissor_number} {current.scissor_type}</label>
          {presetList.length > 0 && (
            <div className="flex gap-1.5 flex-wrap mb-2">
              {presetList.map((p) => {
                const on = presetActive(p);
                return (
                  <button key={p} type="button" onClick={() => togglePreset(p)} title={p}
                    className={`px-2.5 py-1 text-[11px] rounded-lg border text-left leading-snug transition max-w-[300px] truncate ${on ? 'bg-neutral-800 text-white border-neutral-800' : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'}`}>
                    {p.split('\n')[0]}{p.includes('\n') ? ' …' : ''}
                  </button>
                );
              })}
            </div>
          )}
          <p className="text-[11px] text-neutral-400 mb-1.5">※ 장력·밸런스·날각·무뎌짐·찍힘·부품·스토퍼·빗살은 위에서 표시하면 멘트가 자동으로 들어갑니다.</p>
          <textarea value={current.comment} onChange={(e) => patchActive({ comment: e.target.value })} rows={6}
            placeholder="칩 선택·자동삽입된 멘트가 여기 모입니다. 직접 수정도 가능. 줄바꿈 그대로 고객에게 표시됩니다."
            className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm leading-relaxed resize-y focus:outline-none focus:border-neutral-400" />
        </div>
      </div>
    </Card>
  );
}
