'use client';

import { useState } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useSaveInspections } from '@/hooks/use-repairs';
import type { RepairInspection } from '@/lib/supabase/types';
import { Plus, Trash2, Save, Camera, X, Loader2 } from 'lucide-react';
import toast from 'react-hot-toast';
import { resizeImage } from '@/lib/utils/resize-image';
import { createClient } from '@/lib/supabase/client';

/** 검수 항목 선택지 */
const BLADE_OPTIONS = ['양호', '무뎌짐', '찍힘'];
const COMB_OPTIONS = ['', '양호', '손상', '심한손상'];
const TENSION_OPTIONS = ['양호', '헐거움', '교체'];
const PARTS_OPTIONS = ['양호', '교체'];
const STOPPER_OPTIONS = ['양호', '교체'];
const SCISSOR_TYPES = ['블런트', '틴닝', '장가위', '슬라이싱', '기타'];

interface InspectionItem {
  scissor_number: number;
  scissor_type: string;
  blade_tip: string;
  blade_mid: string;
  blade_inner: string;
  comb: string;
  tension: string;
  parts: string;
  stopper: string;
  photo_url: string;
  worker: string;
  /** 로컬 미리보기용 (업로드 전) */
  _photoPreview?: string;
}

function newInspectionItem(num: number): InspectionItem {
  return {
    scissor_number: num,
    scissor_type: '블런트',
    blade_tip: '양호',
    blade_mid: '양호',
    blade_inner: '양호',
    comb: '',
    tension: '양호',
    parts: '양호',
    stopper: '양호',
    photo_url: '',
    worker: '백성민',
  };
}

interface InspectionFormProps {
  repairId: string;
  existingInspections: RepairInspection[];
  totalScissors: number;
  onSaved?: () => void;
}

export function InspectionForm({ repairId, existingInspections, onSaved }: InspectionFormProps) {
  const [items, setItems] = useState<InspectionItem[]>(() => {
    if (existingInspections.length > 0) {
      return existingInspections.map((e) => ({
        scissor_number: e.scissor_number,
        scissor_type: e.scissor_type || '블런트',
        blade_tip: e.blade_tip,
        blade_mid: e.blade_mid,
        blade_inner: e.blade_inner,
        comb: e.comb,
        tension: e.tension,
        parts: e.parts,
        stopper: e.stopper,
        photo_url: e.photo_url || '',
        worker: e.worker,
      }));
    }
    // 기본: 1개만 생성, + 버튼으로 추가
    return [newInspectionItem(1)];
  });
  const [activeIdx, setActiveIdx] = useState(0);

  const saveInspections = useSaveInspections();
  const [uploading, setUploading] = useState(false);

  const updateItem = (idx: number, field: keyof InspectionItem, value: string) => {
    setItems((prev) => {
      const next = [...prev];
      next[idx] = { ...next[idx], [field]: value };
      return next;
    });
  };

  const addScissor = () => {
    const nextNum = items.length + 1;
    setItems((prev) => [...prev, newInspectionItem(nextNum)]);
    setActiveIdx(items.length);
  };

  const removeScissor = (idx: number) => {
    if (items.length <= 1) return;
    setItems((prev) => {
      const next = prev.filter((_, i) => i !== idx);
      return next.map((item, i) => ({ ...item, scissor_number: i + 1 }));
    });
    setActiveIdx(Math.max(0, activeIdx - 1));
  };

  const handlePhotoSelect = async (idx: number, file: File) => {
    // 즉시 미리보기 표시
    const reader = new FileReader();
    reader.onload = (e) => {
      setItems((prev) => {
        const next = [...prev];
        next[idx] = { ...next[idx], _photoPreview: e.target?.result as string };
        return next;
      });
    };
    reader.readAsDataURL(file);

    // 리사이징 후 Storage 업로드
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
      // 업로드 실패 시 미리보기 제거 + 알림
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
      next[idx] = { ...next[idx], photo_url: '', _photoPreview: undefined };
      return next;
    });
  };

  const handleSave = () => {
    // _photoPreview(base64) 제거 — photo_url만 서버 전송
    const cleaned = items.map(({ _photoPreview, ...rest }) => rest);
    saveInspections.mutate({ repairId, inspections: cleaned }, {
      onSuccess: () => { onSaved?.(); },
    });
  };

  const current = items[activeIdx];
  if (!current) return null;

  const isThinning = current.scissor_type === '틴닝';
  const photoSrc = current._photoPreview || current.photo_url;

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center justify-between">
          <span>검수 체크리스트</span>
          <Button
            variant="primary"
            size="sm"
            onClick={handleSave}
            disabled={uploading}
            loading={saveInspections.isPending}
          >
            <Save size={14} />
            {uploading ? '사진 업로드 중...' : '저장'}
          </Button>
        </CardTitle>
      </CardHeader>

      {/* 가위 탭 — + 버튼으로 추가 */}
      <div className="flex gap-1 mb-4 overflow-x-auto">
        {items.map((item, idx) => (
          <button
            key={idx}
            onClick={() => setActiveIdx(idx)}
            className={`shrink-0 px-3 py-1.5 text-xs font-medium rounded-lg transition ${
              activeIdx === idx
                ? 'bg-terracotta text-white'
                : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'
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

      {/* 체크리스트 */}
      <div className="space-y-3">
        {/* 가위 종류 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0">종류</label>
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

        {/* 사진 촬영/업로드 — 종류 바로 밑 */}
        <div className="flex items-start gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0 pt-2">사진</label>
          {photoSrc ? (
            <div className="relative">
              <img src={photoSrc} alt="검수 사진" className="w-28 h-28 object-cover rounded-lg border border-neutral-200" />
              {/* 업로드 중 오버레이 */}
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
              {/* 재촬영 */}
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
            <label className="cursor-pointer flex flex-col items-center justify-center w-28 h-28 border-2 border-dashed border-neutral-300 rounded-lg hover:border-terracotta/50 transition">
              {uploading ? (
                <Loader2 size={24} className="text-terracotta mb-1 animate-spin" />
              ) : (
                <Camera size={24} className="text-neutral-400 mb-1" />
              )}
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

        {/* 날 상태 */}
        {(['blade_tip', 'blade_mid', 'blade_inner'] as const).map((field) => {
          const label = field === 'blade_tip' ? '날끝' : field === 'blade_mid' ? '날중간' : '날안쪽';
          return (
            <div key={field} className="flex items-center gap-3">
              <label className="text-sm text-neutral-500 w-20 shrink-0">{label}</label>
              <div className="flex gap-1.5 flex-wrap">
                {BLADE_OPTIONS.map((opt) => (
                  <button
                    key={opt}
                    onClick={() => updateItem(activeIdx, field, opt)}
                    className={`px-3 py-1.5 text-xs rounded-lg border transition ${
                      current[field] === opt
                        ? opt === '양호' ? 'bg-success-soft text-success border-success/30'
                          : 'bg-error-soft text-error border-error/30'
                        : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'
                    }`}
                  >
                    {opt}
                  </button>
                ))}
              </div>
            </div>
          );
        })}

        {/* 빗살 (틴닝 전용) */}
        {isThinning && (
          <div className="flex items-center gap-3">
            <label className="text-sm text-neutral-500 w-20 shrink-0">빗살</label>
            <div className="flex gap-1.5 flex-wrap">
              {COMB_OPTIONS.filter(Boolean).map((opt) => (
                <button
                  key={opt}
                  onClick={() => updateItem(activeIdx, 'comb', opt)}
                  className={`px-3 py-1.5 text-xs rounded-lg border transition ${
                    current.comb === opt
                      ? opt === '양호' ? 'bg-success-soft text-success border-success/30'
                        : 'bg-error-soft text-error border-error/30'
                      : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'
                  }`}
                >
                  {opt}
                </button>
              ))}
            </div>
          </div>
        )}

        {/* 장력 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0">장력</label>
          <div className="flex gap-1.5">
            {TENSION_OPTIONS.map((opt) => (
              <button
                key={opt}
                onClick={() => updateItem(activeIdx, 'tension', opt)}
                className={`px-3 py-1.5 text-xs rounded-lg border transition ${
                  current.tension === opt
                    ? opt === '양호' ? 'bg-success-soft text-success border-success/30'
                      : 'bg-error-soft text-error border-error/30'
                    : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'
                }`}
              >
                {opt}
              </button>
            ))}
          </div>
        </div>

        {/* 내부부품 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0">내부부품</label>
          <div className="flex gap-1.5">
            {PARTS_OPTIONS.map((opt) => (
              <button
                key={opt}
                onClick={() => updateItem(activeIdx, 'parts', opt)}
                className={`px-3 py-1.5 text-xs rounded-lg border transition ${
                  current.parts === opt
                    ? opt === '양호' ? 'bg-success-soft text-success border-success/30'
                      : 'bg-warning-soft text-warning border-warning/30'
                    : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'
                }`}
              >
                {opt}
              </button>
            ))}
          </div>
        </div>

        {/* 스토퍼 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0">스토퍼</label>
          <div className="flex gap-1.5">
            {STOPPER_OPTIONS.map((opt) => (
              <button
                key={opt}
                onClick={() => updateItem(activeIdx, 'stopper', opt)}
                className={`px-3 py-1.5 text-xs rounded-lg border transition ${
                  current.stopper === opt
                    ? opt === '양호' ? 'bg-success-soft text-success border-success/30'
                      : 'bg-warning-soft text-warning border-warning/30'
                    : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'
                }`}
              >
                {opt}
              </button>
            ))}
          </div>
        </div>

        {/* 작업자 */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0">작업자</label>
          <input
            type="text"
            value={current.worker}
            onChange={(e) => updateItem(activeIdx, 'worker', e.target.value)}
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm"
          />
        </div>
      </div>
    </Card>
  );
}
