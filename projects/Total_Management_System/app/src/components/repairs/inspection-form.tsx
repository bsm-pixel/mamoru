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
import { InspectionPhotoMarker, type PhotoMark } from './inspection-photo-marker';
import { COMMENT_PRESETS } from '@/lib/repair/comment-presets';

const SCISSOR_TYPES = ['블런트', '틴닝', '장가위', '슬라이싱', '기타'];

interface InspectionItem {
  scissor_number: number;
  scissor_type: string;
  photo_url: string;
  photo_marks: PhotoMark[];
  worker: string;
  /** 로컬 미리보기용 (업로드 전) */
  _photoPreview?: string;
}

function newInspectionItem(num: number): InspectionItem {
  return {
    scissor_number: num,
    scissor_type: '블런트',
    photo_url: '',
    photo_marks: [],
    worker: '백성민',
  };
}

interface InspectionFormProps {
  repairId: string;
  existingInspections: RepairInspection[];
  totalScissors: number;
  /** 고객 노출 코멘트 초기값 (repairs.admin_note) */
  initialComment?: string;
  onSaved?: () => void;
  /** 디자인 모니터 데모 모드 — 스토리지 업로드/API 저장/draft 모두 비활성 (로컬 미리보기만) */
  demo?: boolean;
}

/** 모바일 가로 감지 (검수 화면 세로 고정용) */
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

function draftKey(repairId: string) {
  return `repair_inspect_${repairId}`;
}

interface Draft {
  items: InspectionItem[];
  comment: string;
  activeIdx: number;
}

function loadDraft(repairId: string): Draft | null {
  if (typeof window === 'undefined') return null;
  try {
    const raw = sessionStorage.getItem(draftKey(repairId));
    return raw ? (JSON.parse(raw) as Draft) : null;
  } catch {
    return null;
  }
}

export function InspectionForm({ repairId, existingInspections, initialComment, onSaved, demo = false }: InspectionFormProps) {
  const draft = !demo && typeof window !== 'undefined' ? loadDraft(repairId) : null;

  const [items, setItems] = useState<InspectionItem[]>(() => {
    if (draft?.items?.length) return draft.items;
    if (existingInspections.length > 0) {
      return existingInspections.map((e) => ({
        scissor_number: e.scissor_number,
        scissor_type: e.scissor_type || '블런트',
        photo_url: e.photo_url || '',
        photo_marks: (e.photo_marks as PhotoMark[] | null) || [],
        worker: e.worker,
      }));
    }
    return [newInspectionItem(1)];
  });
  const [activeIdx, setActiveIdx] = useState(draft?.activeIdx ?? 0);
  const [comment, setComment] = useState(draft?.comment ?? initialComment ?? '');
  const [uploading, setUploading] = useState(false);
  const [savingComment, setSavingComment] = useState(false);

  const saveInspections = useSaveInspections();
  const landscapeMobile = useLandscapeMobile();

  // 작성 중 draft 보존 (카메라/포커스 손실로 모달 재마운트 시 복원)
  useEffect(() => {
    if (demo || typeof window === 'undefined') return;
    try {
      sessionStorage.setItem(draftKey(repairId), JSON.stringify({ items, comment, activeIdx }));
    } catch { /* quota 초과 등 무시 */ }
  }, [demo, repairId, items, comment, activeIdx]);

  const updateItem = (idx: number, field: 'scissor_type' | 'worker', value: string) => {
    setItems((prev) => {
      const next = [...prev];
      next[idx] = { ...next[idx], [field]: value };
      return next;
    });
  };

  const updateMarks = (idx: number, marks: PhotoMark[]) => {
    setItems((prev) => {
      const next = [...prev];
      next[idx] = { ...next[idx], photo_marks: marks };
      return next;
    });
  };

  const addScissor = () => {
    setItems((prev) => [...prev, newInspectionItem(prev.length + 1)]);
    setActiveIdx(items.length);
  };

  const removeScissor = (idx: number) => {
    if (items.length <= 1) return;
    setItems((prev) => prev.filter((_, i) => i !== idx).map((item, i) => ({ ...item, scissor_number: i + 1 })));
    setActiveIdx(Math.max(0, activeIdx - 1));
  };

  const handlePhotoSelect = async (idx: number, file: File) => {
    // 즉시 미리보기
    const reader = new FileReader();
    reader.onload = (e) => {
      const url = e.target?.result as string;
      setItems((prev) => {
        const next = [...prev];
        // 데모: 로컬 data URL 을 photo_url 로 바로 사용(스토리지 업로드 없이 핀 마킹 가능)
        next[idx] = { ...next[idx], _photoPreview: url, ...(demo ? { photo_url: url } : {}) };
        return next;
      });
    };
    reader.readAsDataURL(file);

    if (demo) return; // 데모: 스토리지 업로드 생략

    setUploading(true);
    try {
      const resized = await resizeImage(file, 1200, 0.8);
      const supabase = createClient();
      const ext = file.name.split('.').pop() || 'jpg';
      const filePath = `inspections/${repairId}/${Date.now()}_${idx}.${ext}`;

      const { error: uploadErr } = await supabase.storage
        .from('repair-photos')
        .upload(filePath, resized, { contentType: 'image/jpeg', upsert: false });

      if (uploadErr) throw uploadErr;

      const { data: urlData } = supabase.storage.from('repair-photos').getPublicUrl(filePath);

      setItems((prev) => {
        const next = [...prev];
        next[idx] = { ...next[idx], photo_url: urlData.publicUrl };
        return next;
      });
    } catch (err) {
      console.error('검수 사진 업로드 실패:', err);
      setItems((prev) => {
        const next = [...prev];
        next[idx] = { ...next[idx], _photoPreview: undefined, photo_url: '' };
        return next;
      });
      toast.error('사진 업로드 실패 — 다시 시도해주세요');
    } finally {
      setUploading(false);
    }
  };

  const removePhoto = (idx: number) => {
    setItems((prev) => {
      const next = [...prev];
      next[idx] = { ...next[idx], photo_url: '', _photoPreview: undefined, photo_marks: [] };
      return next;
    });
  };

  const togglePreset = (phrase: string) => {
    setComment((prev) => {
      const lines = prev.split('\n').map((l) => l.trim()).filter(Boolean);
      const idx = lines.indexOf(phrase);
      if (idx >= 0) lines.splice(idx, 1);
      else lines.push(phrase);
      return lines.join('\n');
    });
  };
  const presetActive = (phrase: string) =>
    comment.split('\n').map((l) => l.trim()).includes(phrase);

  const handleSave = async () => {
    if (demo) {
      toast('데모 화면입니다 — 실제로 저장되지 않습니다', { icon: '🎨' });
      return;
    }
    // 1) 고객 노출 코멘트(admin_note) 저장
    setSavingComment(true);
    try {
      await fetch(`/api/repair/${repairId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ admin_note: comment.trim() || null }),
      });
    } catch (err) {
      console.error('코멘트 저장 실패:', err);
    } finally {
      setSavingComment(false);
    }

    // 2) 검수(가위별 사진+핀) 저장 — _photoPreview 제거
    const cleaned = items.map(({ _photoPreview, ...rest }) => rest);
    saveInspections.mutate(
      { repairId, inspections: cleaned },
      {
        onSuccess: () => {
          try { sessionStorage.removeItem(draftKey(repairId)); } catch { /* noop */ }
          onSaved?.();
        },
      }
    );
  };

  const current = items[activeIdx];
  if (!current) return null;

  const photoSrc = current._photoPreview || current.photo_url;
  const busy = uploading || savingComment || saveInspections.isPending;

  return (
    <Card>
      {/* 모바일 가로 차단 오버레이 (검수는 세로 작업) */}
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
            {uploading ? '사진 업로드 중...' : savingComment ? '저장 중...' : '저장'}
          </Button>
        </CardTitle>
      </CardHeader>

      {/* 가위 탭 */}
      <div className="flex gap-1 mb-4 overflow-x-auto">
        {items.map((item, idx) => (
          <button
            key={idx}
            onClick={() => setActiveIdx(idx)}
            className={`shrink-0 px-3 py-1.5 text-xs font-medium rounded-lg transition ${
              activeIdx === idx ? 'bg-neutral-800 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
            }`}
          >
            #{item.scissor_number} {item.scissor_type}
          </button>
        ))}
        <button
          onClick={addScissor}
          className="shrink-0 px-3 py-1.5 text-xs text-neutral-400 hover:text-neutral-600 flex items-center gap-1"
        >
          <Plus size={12} /> 추가
        </button>
      </div>

      <div className="space-y-4">
        {/* 종류 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0">종류</label>
          <select
            value={current.scissor_type}
            onChange={(e) => updateItem(activeIdx, 'scissor_type', e.target.value)}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm"
          >
            {SCISSOR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          {items.length > 1 && (
            <button onClick={() => removeScissor(activeIdx)} className="text-error/60 hover:text-error">
              <Trash2 size={14} />
            </button>
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
                <div className="absolute inset-0 bg-black/40 rounded-lg flex items-center justify-center">
                  <Loader2 size={24} className="text-white animate-spin" />
                </div>
              )}
              <button
                onClick={() => removePhoto(activeIdx)}
                className="absolute -top-2 -right-2 w-7 h-7 bg-error text-white rounded-full flex items-center justify-center shadow-md"
              >
                <X size={14} />
              </button>
              <label className="absolute bottom-1 right-1 w-7 h-7 bg-white/80 rounded-full flex items-center justify-center cursor-pointer shadow hover:bg-white transition">
                <Camera size={14} className="text-neutral-600" />
                <input
                  type="file"
                  accept="image/*"
                  capture="environment"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) handlePhotoSelect(activeIdx, file);
                    e.target.value = '';
                  }}
                  className="hidden"
                />
              </label>
            </div>
          ) : (
            <label className="cursor-pointer flex flex-col items-center justify-center w-28 h-28 border-2 border-dashed border-neutral-300 rounded-lg hover:border-neutral-400 transition">
              {uploading ? <Loader2 size={24} className="text-neutral-500 mb-1 animate-spin" /> : <Camera size={24} className="text-neutral-400 mb-1" />}
              <span className="text-[11px] text-neutral-400">{uploading ? '업로드중...' : '촬영/업로드'}</span>
              <input
                type="file"
                accept="image/*"
                capture="environment"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) handlePhotoSelect(activeIdx, file);
                  e.target.value = '';
                }}
                className="hidden"
              />
            </label>
          )}
        </div>

        {/* 핀 마킹 — 사진 위 상처 표시 (체크리스트 대체) */}
        <div>
          <label className="block text-sm text-neutral-500 mb-2">상처 표시 (사진 위 핀)</label>
          <InspectionPhotoMarker
            photoUrl={current.photo_url}
            marks={current.photo_marks}
            onChange={(marks) => updateMarks(activeIdx, marks)}
          />
        </div>

        {/* 작업자 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-16 shrink-0">작업자</label>
          <input
            type="text"
            value={current.worker}
            onChange={(e) => updateItem(activeIdx, 'worker', e.target.value)}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm"
          />
        </div>

        {/* 진단 멘트 — 프리셋 체크 → 줄바꿈 삽입 + 직접 수정 (고객 노출, 가위 공통) */}
        <div className="pt-3 border-t border-neutral-100">
          <label className="block text-sm font-medium text-neutral-600 mb-2">진단 멘트 (고객 안내문)</label>
          <div className="flex gap-1.5 flex-wrap mb-2">
            {COMMENT_PRESETS.map((p) => {
              const active = presetActive(p);
              return (
                <button
                  key={p}
                  type="button"
                  onClick={() => togglePreset(p)}
                  className={`px-2.5 py-1 text-[11px] rounded-lg border text-left leading-snug transition max-w-full ${
                    active
                      ? 'bg-neutral-800 text-white border-neutral-800'
                      : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'
                  }`}
                >
                  {p}
                </button>
              );
            })}
          </div>
          <textarea
            value={comment}
            onChange={(e) => setComment(e.target.value)}
            rows={5}
            placeholder="위 문구를 선택하거나 직접 작성하세요. 줄바꿈 그대로 고객에게 표시됩니다."
            className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm leading-relaxed resize-y focus:outline-none focus:border-neutral-400"
          />
        </div>
      </div>
    </Card>
  );
}
