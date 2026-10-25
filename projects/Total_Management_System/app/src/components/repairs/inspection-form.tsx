'use client';

import { useState, useEffect, useRef, type ReactNode } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useSaveInspections } from '@/hooks/use-repairs';
import type { RepairInspection } from '@/lib/supabase/types';
import { Plus, Trash2, Save, Camera, X, Loader2, ImagePlus, ChevronDown } from 'lucide-react';
import toast from 'react-hot-toast';
import { compressImage } from '@/lib/utils/compress-image';
import { createClient } from '@/lib/supabase/client';
import { CameraCapture } from '@/components/ui/camera-capture';
import { COMMENT_PRESETS_BY_TYPE } from '@/lib/repair/comment-presets';
import { SEGS, SEG_KO, kindsFor, checkIdsOf, normalizeChecks, woundLines } from '@/lib/repair/wound-catalog';
import { WoundDiagram } from './wound-diagram';

const SCISSOR_TYPES = ['블런트', '틴닝', '장가위', '슬라이싱', '기타'];
const MAX_ZOOM = 3;

/**
 * 수리내역서 v3 (2026-10-15) — 사진 위 손 핀 → 가위 종류별 그림체 + 상처 체크
 *   가위 사진 1장(1:1) · 확대 사진 0~3장 · 상처 칩 체크 · 추가 메모(선택)
 *   체크 항목·문구 SSOT = lib/repair/wound-catalog.ts
 *
 * 구 내역서(diagram_marks = NULL, 사진 핀)는 열어도 그대로 보존된다.
 *   상처를 하나라도 체크해야 새 방식으로 바뀐다 (그 전엔 저장해도 고객 화면이 옛 방식 그대로)
 */
interface InspectionItem {
  scissor_number: number;
  scissor_type: string;
  photo_url: string;
  zoom_photos: string[];
  checks: string[];
  /** v2 사진 핀 내역서 — 체크 전까지 diagram_marks 를 NULL 로 둔다 */
  legacy: boolean;
  /** v2 핀 데이터 비파괴 통과 */
  photo_marks: RepairInspection['photo_marks'];
  comment: string;
  worker: string;
}

function newInspectionItem(num: number): InspectionItem {
  return { scissor_number: num, scissor_type: '블런트', photo_url: '', zoom_photos: [], checks: [], legacy: false, photo_marks: null, comment: '', worker: '백성민' };
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
// base = 초안을 만들 때 기준이 된 DB 내역서(행 id 목록). 저장할 때마다 행이 새로 만들어져 id가 바뀐다.
// v = 초안 형식 버전. v3 이전 초안(marks/flags 형식)은 버린다
interface Draft { v: 3; items: InspectionItem[]; activeIdx: number; base?: string; }
const baseSig = (rows: RepairInspection[]) => rows.map((r) => r.id).join(',');
function loadDraft(repairId: string, base: string): Draft | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = sessionStorage.getItem(draftKey(repairId));
    if (!raw) return null;
    const d = JSON.parse(raw) as Draft;
    // 초안 이후 DB가 바뀌었거나(다른 기기 저장 등) 옛 형식이면 버린다 — 오래된 초안이 최신 내역서를 덮지 않게
    if (d.v !== 3 || d.base !== base) { sessionStorage.removeItem(draftKey(repairId)); return null; }
    return d;
  } catch { return null; }
}

async function uploadRepairPhoto(repairId: string, file: File, tag: string): Promise<string> {
  const supabase = createClient();
  const ext = file.type === 'image/webp' ? 'webp' : 'jpg';
  const filePath = `inspections/${repairId}/${Date.now()}_${tag}_${Math.random().toString(36).slice(2, 7)}.${ext}`;
  const { error } = await supabase.storage.from('repair-photos').upload(filePath, file, { contentType: file.type, upsert: false });
  if (error) throw error;
  return supabase.storage.from('repair-photos').getPublicUrl(filePath).data.publicUrl;
}

export function InspectionForm({ repairId, existingInspections, totalScissors, onSaved }: InspectionFormProps) {
  const base = baseSig(existingInspections);
  const draft = typeof window !== 'undefined' ? loadDraft(repairId, base) : null;

  const [items, setItems] = useState<InspectionItem[]>(() => {
    if (draft?.items?.length) return draft.items;
    if (existingInspections.length > 0) {
      return existingInspections.map((e) => ({
        scissor_number: e.scissor_number,
        scissor_type: e.scissor_type || '블런트',
        photo_url: e.photo_url || '',
        zoom_photos: Array.isArray(e.zoom_photos) ? e.zoom_photos : [],
        checks: normalizeChecks(e.diagram_marks),
        legacy: e.diagram_marks == null,
        photo_marks: e.photo_marks,
        comment: e.comment || '',
        worker: e.worker,
      }));
    }
    // 접수 수량(마모루+타사)만큼 가위 탭을 미리 만든다
    return Array.from({ length: Math.max(1, totalScissors || 0) }, (_, i) => newInspectionItem(i + 1));
  });
  const [activeIdx, setActiveIdx] = useState(draft?.activeIdx ?? 0);
  // 업로드 중 표시 — 미리보기는 메모리(object URL)에만. 초안(sessionStorage)엔 넣지 않는다(수 MB 라 저장 실패 → 초안 전체 유실)
  const [pending, setPending] = useState<Record<string, string>>({});
  const [busyCount, setBusyCount] = useState(0);
  const [cameraOpen, setCameraOpen] = useState(false);
  const [memoOpen, setMemoOpen] = useState(false);
  const zoomInput = useRef<HTMLInputElement>(null);

  const saveInspections = useSaveInspections();
  const landscapeMobile = useLandscapeMobile();

  useEffect(() => {
    if (typeof window === 'undefined') return;
    try { sessionStorage.setItem(draftKey(repairId), JSON.stringify({ v: 3, items, activeIdx, base })); } catch { /* noop */ }
  }, [repairId, items, activeIdx, base]);

  const patchAt = (idx: number, fn: (it: InspectionItem) => InspectionItem) =>
    setItems((prev) => prev.map((it, i) => (i === idx ? fn(it) : it)));
  const patchActive = (patch: Partial<InspectionItem>) => patchAt(activeIdx, (it) => ({ ...it, ...patch }));

  const addScissor = () => { setItems((prev) => [...prev, newInspectionItem(prev.length + 1)]); setActiveIdx(items.length); };
  const removeScissor = (idx: number) => {
    if (items.length <= 1) return;
    // 중간 가위를 지우면 뒤 번호가 당겨져 이미 붙인 꼬리표(01/03 등)와 어긋난다 → 확인 후 진행
    if (idx < items.length - 1 && !window.confirm(`#${idx + 1} 가위를 지우면 뒤 가위 번호가 하나씩 당겨집니다.\n꼬리표를 이미 붙였다면 지운 뒤 꼬리표를 다시 출력해 주세요.\n\n지울까요?`)) return;
    setItems((prev) => prev.filter((_, i) => i !== idx).map((it, i) => ({ ...it, scissor_number: i + 1 })));
    setActiveIdx(Math.max(0, activeIdx - 1));
  };

  // 종류 변경 — 새 종류에 없는 항목(예: 틴닝 전용 빗살)은 체크 해제
  const changeType = (t: string) => patchAt(activeIdx, (it) => {
    const allowed = new Set(kindsFor(t).flatMap(checkIdsOf));
    return { ...it, scissor_type: t, checks: it.checks.filter((c) => allowed.has(c)) };
  });

  // 체크 토글 — 구 내역서도 체크하는 순간 새 방식(legacy 해제)
  const toggleCheck = (id: string) => patchAt(activeIdx, (it) => ({
    ...it,
    legacy: false,
    checks: it.checks.includes(id) ? it.checks.filter((c) => c !== id) : [...it.checks, id],
  }));

  const runUpload = async (key: string, file: File, opts: Parameters<typeof compressImage>[1], tag: string) => {
    setPending((p) => ({ ...p, [key]: URL.createObjectURL(file) }));
    setBusyCount((n) => n + 1);
    try {
      const small = await compressImage(file, opts);
      return await uploadRepairPhoto(repairId, small, tag);
    } finally {
      setPending((p) => { const n = { ...p }; if (n[key]) URL.revokeObjectURL(n[key]); delete n[key]; return n; });
      setBusyCount((n) => n - 1);
    }
  };

  // 가위 사진 — 1:1 가운데 자르기, 1080. 재촬영해도 체크는 유지(그림체는 사진과 무관)
  const handlePhotoSelect = async (idx: number, file: File) => {
    try {
      const url = await runUpload(`main${idx}`, file, { square: true, maxSize: 1080 }, `s${idx + 1}`);
      patchAt(idx, (it) => ({ ...it, photo_url: url }));
    } catch (err) {
      console.error('가위 사진 업로드 실패:', err);
      toast.error('사진 업로드 실패 — 다시 시도해주세요');
    }
  };

  // 확대 사진 — 갤러리 여러 장, 긴 변 1080 (현미경 원본 비율 유지)
  const handleZoomFiles = async (idx: number, files: File[]) => {
    const room = MAX_ZOOM - items[idx].zoom_photos.length;
    if (files.length > room) toast(`확대 사진은 가위당 ${MAX_ZOOM}장까지라 ${room}장만 올립니다`);
    await Promise.all(files.slice(0, Math.max(0, room)).map(async (f, k) => {
      try {
        const url = await runUpload(`zoom${idx}_${Date.now()}_${k}`, f, { maxSize: 1080 }, `z${idx + 1}`);
        patchAt(idx, (it) => ({ ...it, zoom_photos: [...it.zoom_photos, url].slice(0, MAX_ZOOM) }));
      } catch (err) {
        console.error('확대 사진 업로드 실패:', err);
        toast.error('확대 사진 업로드 실패 — 다시 시도해주세요');
      }
    }));
  };

  const togglePreset = (p: string) => {
    const t = p.trim();
    patchAt(activeIdx, (it) => {
      const blocks = it.comment.split('\n\n').map((b) => b.trim()).filter(Boolean);
      const k = blocks.indexOf(t);
      if (k >= 0) blocks.splice(k, 1); else blocks.push(t);
      return { ...it, comment: blocks.join('\n\n') };
    });
  };

  const handleSave = () => {
    const rows = items.map((it) => ({
      scissor_number: it.scissor_number,
      scissor_type: it.scissor_type,
      photo_url: it.photo_url || null,
      photo_marks: it.photo_marks ?? null,
      // 구 내역서를 체크 없이 저장 → NULL 유지(고객 화면 옛 방식 그대로)
      diagram_marks: it.legacy && it.checks.length === 0 ? null : it.checks,
      zoom_photos: it.zoom_photos,
      comment: it.comment.trim() || null,
      worker: it.worker,
    }));
    saveInspections.mutate({ repairId, inspections: rows }, {
      onSuccess: () => { try { sessionStorage.removeItem(draftKey(repairId)); } catch { /* noop */ } onSaved?.(); },
    });
  };

  const current = items[activeIdx];
  if (!current) return null;
  const mainPending = pending[`main${activeIdx}`];
  const photoSrc = mainPending || current.photo_url;
  const zoomPending = Object.entries(pending).filter(([k]) => k.startsWith(`zoom${activeIdx}_`)).map(([, v]) => v);
  const busy = busyCount > 0 || saveInspections.isPending;
  const kinds = kindsFor(current.scissor_type);
  const segKinds = kinds.filter((k) => k.segs);
  const plainKinds = kinds.filter((k) => !k.segs);
  const lines = woundLines(current.checks);
  const presetList = COMMENT_PRESETS_BY_TYPE[current.scissor_type] || [];
  const presetActive = (p: string) => current.comment.split('\n\n').map((b) => b.trim()).includes(p.trim());

  const chip = (id: string, label: ReactNode) => {
    const on = current.checks.includes(id);
    return (
      <button key={id} type="button" onClick={() => toggleCheck(id)}
        className={`min-h-9 px-3 rounded-md border text-[13px] font-medium transition inline-flex items-center gap-1.5 ${on ? 'bg-neutral-900 border-neutral-900 text-white' : 'bg-white border-neutral-200 text-neutral-600 hover:border-neutral-300'}`}>
        {label}
      </button>
    );
  };
  const sw = (color: string) => <i className="inline-block w-2.5 h-2.5 rounded-[2px] shrink-0" style={{ background: color }} />;

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
        <CardTitle className="flex items-center justify-between w-full">
          <span>가위 {items.length}자루 · 사진 {items.filter((it) => it.photo_url).length}/{items.length}</span>
          <Button variant="primary" size="sm" onClick={handleSave} disabled={busy} loading={saveInspections.isPending}>
            <Save size={14} />
            {busyCount > 0 ? '사진 올리는 중...' : '저장'}
          </Button>
        </CardTitle>
      </CardHeader>

      {/* 가위 탭 */}
      <div className="flex gap-1 mb-4 overflow-x-auto">
        {items.map((item, idx) => (
          <button key={idx} onClick={() => setActiveIdx(idx)}
            className={`shrink-0 px-3 py-1.5 text-xs font-medium rounded-lg transition ${activeIdx === idx ? 'bg-neutral-800 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'}`}>
            #{item.scissor_number} {item.scissor_type}{item.photo_url ? '' : ' ·'}
          </button>
        ))}
        <button onClick={addScissor} className="shrink-0 px-3 py-1.5 text-xs text-neutral-400 hover:text-neutral-600 flex items-center gap-1">
          <Plus size={12} /> 추가
        </button>
      </div>

      <div className="space-y-5">
        {/* 종류 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0">종류</label>
          <select value={current.scissor_type} onChange={(e) => changeType(e.target.value)}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm">
            {SCISSOR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          {items.length > 1 && (
            <button onClick={() => removeScissor(activeIdx)} title="이 가위 삭제" className="text-error/60 hover:text-error"><Trash2 size={14} /></button>
          )}
        </div>

        {/* 사진 — 가위 1장(1:1) + 확대 사진 0~3장 */}
        <div className="flex items-start gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0 pt-2">사진</label>
          <div className="flex flex-wrap items-start gap-2">
            {photoSrc ? (
              <div className="relative">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img src={photoSrc} alt="가위 사진" className="w-28 h-28 object-cover rounded-lg border border-neutral-200" />
                {mainPending && <div className="absolute inset-0 bg-black/40 rounded-lg flex items-center justify-center"><Loader2 size={24} className="text-white animate-spin" /></div>}
                <button type="button" onClick={() => setCameraOpen(true)} title="다시 찍기"
                  className="absolute bottom-1 right-1 w-7 h-7 bg-white/85 rounded-full flex items-center justify-center shadow hover:bg-white transition">
                  <Camera size={14} className="text-neutral-600" />
                </button>
              </div>
            ) : (
              <button type="button" onClick={() => setCameraOpen(true)}
                className="flex flex-col items-center justify-center w-28 h-28 border-2 border-dashed border-neutral-300 rounded-lg hover:border-neutral-400 transition">
                <Camera size={24} className="text-neutral-400 mb-1" />
                <span className="text-[11px] text-neutral-400">가위 사진</span>
              </button>
            )}

            {/* 확대 사진 */}
            {current.zoom_photos.map((u, k) => (
              <div key={u} className="relative">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img src={u} alt={`확대 사진 ${k + 1}`} className="w-16 h-16 object-cover rounded-md border border-neutral-200" />
                <button type="button" onClick={() => patchActive({ zoom_photos: current.zoom_photos.filter((x) => x !== u) })} title="빼기"
                  className="absolute -top-1.5 -right-1.5 w-5 h-5 bg-neutral-800 text-white rounded-full flex items-center justify-center"><X size={11} /></button>
              </div>
            ))}
            {zoomPending.map((u) => (
              <div key={u} className="relative">
                {/* eslint-disable-next-line @next/next/no-img-element */}
                <img src={u} alt="" className="w-16 h-16 object-cover rounded-md border border-neutral-200 opacity-60" />
                <Loader2 size={18} className="absolute inset-0 m-auto text-neutral-700 animate-spin" />
              </div>
            ))}
            {current.zoom_photos.length + zoomPending.length < MAX_ZOOM && (
              <button type="button" onClick={() => zoomInput.current?.click()}
                className="w-16 h-16 flex flex-col items-center justify-center border border-dashed border-neutral-300 rounded-md text-neutral-400 hover:border-neutral-400 transition">
                <ImagePlus size={16} />
                <span className="text-[10px] mt-0.5">확대 사진</span>
              </button>
            )}
            {/* capture 미사용 — 갤러리 선택(현미경 앱으로 찍어둔 사진) */}
            <input ref={zoomInput} type="file" accept="image/*" multiple className="hidden"
              onChange={(e) => { const fs = Array.from(e.target.files || []); e.target.value = ''; if (fs.length) handleZoomFiles(activeIdx, fs); }} />
          </div>
        </div>

        <CameraCapture open={cameraOpen} square hiRes onClose={() => setCameraOpen(false)} onCapture={(file) => handlePhotoSelect(activeIdx, file)} />

        {/* 상처 체크 */}
        <div className="pt-4 border-t border-neutral-100">
          <label className="block text-sm font-medium text-neutral-600 mb-2">상처 체크 — #{current.scissor_number} {current.scissor_type}</label>
          {current.legacy && (
            <p className="text-[12px] text-amber-700 bg-amber-50 rounded-md px-2.5 py-2 mb-3">
              이전 방식(사진 위 표시)으로 작성된 내역입니다. 상처를 체크하면 새 방식으로 바뀝니다. 체크하지 않고 저장하면 고객 화면은 그대로입니다.
            </p>
          )}
          <div className="space-y-2">
            {segKinds.map((k) => (
              <div key={k.key} className="flex items-center gap-2 flex-wrap">
                <span className="w-16 shrink-0 text-[13px] font-semibold flex items-center gap-1.5">{sw(k.color)}{k.ko}</span>
                {SEGS.map((s) => chip(`z-${s}-${k.key}`, SEG_KO[s]))}
              </div>
            ))}
            <div className="flex gap-2 flex-wrap pt-1">
              {plainKinds.map((k) => chip(`z-${k.key}`, <>{sw(k.color)}{k.ko}</>))}
            </div>
          </div>

          {/* 고객 화면 미리보기 — 그림 + 목록 (고객 페이지와 같은 규칙) */}
          {lines.length > 0 ? (
            <div className="mt-4 space-y-3">
              <WoundDiagram scissorType={current.scissor_type} checks={current.checks} />
              <ul className="divide-y divide-neutral-100 border-y border-neutral-100">
                {lines.map((l) => (
                  <li key={l.key} className="flex gap-2.5 py-2">
                    <i className="mt-1.5 w-2.5 h-2.5 rounded-[2px] shrink-0" style={{ background: l.color }} />
                    <div>
                      <p className="text-sm font-semibold">{l.title}</p>
                      {l.desc.map((d, i) => <p key={i} className={`text-xs ${i ? 'text-neutral-400 mt-0.5' : 'text-neutral-600'}`}>{d}</p>)}
                    </div>
                  </li>
                ))}
              </ul>
            </div>
          ) : !current.legacy && (
            <p className="mt-3 text-xs text-neutral-400">체크한 상처가 없으면 고객에게 &quot;특별한 상처 없이 양호&quot;로 표시됩니다.</p>
          )}
        </div>

        {/* 작업자 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0">작업자</label>
          <input type="text" value={current.worker} onChange={(e) => patchActive({ worker: e.target.value })}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
        </div>

        {/* 추가 메모 (선택) — 상처 목록 외에 고객에게 더 전할 말. 기본 접힘 */}
        <div className="pt-3 border-t border-neutral-100">
          <button type="button" onClick={() => setMemoOpen((v) => !v)} className="flex items-center gap-1 text-sm text-neutral-500">
            <ChevronDown size={14} className={`transition ${memoOpen || current.comment ? 'rotate-180' : ''}`} />
            추가 안내 (선택){current.comment ? ' · 작성됨' : ''}
          </button>
          {(memoOpen || !!current.comment) && (
            <div className="mt-2">
              {presetList.length > 0 && (
                <div className="flex gap-1.5 flex-wrap mb-2">
                  {presetList.map((p) => (
                    <button key={p} type="button" onClick={() => togglePreset(p)} title={p}
                      className={`px-2.5 py-1 text-[11px] rounded-lg border text-left leading-snug transition max-w-[300px] truncate ${presetActive(p) ? 'bg-neutral-800 text-white border-neutral-800' : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'}`}>
                      {p.split('\n')[0]}{p.includes('\n') ? ' …' : ''}
                    </button>
                  ))}
                </div>
              )}
              <textarea value={current.comment} onChange={(e) => patchActive({ comment: e.target.value })} rows={4}
                placeholder="상처 목록 외에 고객에게 더 전할 내용 (줄바꿈 그대로 표시)"
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm leading-relaxed resize-y focus:outline-none focus:border-neutral-400" />
            </div>
          )}
        </div>
      </div>
    </Card>
  );
}
