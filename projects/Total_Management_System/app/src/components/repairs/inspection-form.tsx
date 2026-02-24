'use client';

import { useState } from 'react';
import { Card, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { useSaveInspections } from '@/hooks/use-repairs';
import type { RepairInspection } from '@/lib/supabase/types';
import { Plus, Trash2, Save } from 'lucide-react';

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
}

export function InspectionForm({ repairId, existingInspections, totalScissors }: InspectionFormProps) {
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
    // 기본: totalScissors 수만큼 빈 항목
    return Array.from({ length: Math.max(totalScissors, 1) }, (_, i) => newInspectionItem(i + 1));
  });
  const [activeIdx, setActiveIdx] = useState(0);

  const saveInspections = useSaveInspections();

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

  const handleSave = () => {
    saveInspections.mutate({ repairId, inspections: items });
  };

  const current = items[activeIdx];
  if (!current) return null;

  const isThinning = current.scissor_type === '틴닝';

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center justify-between">
          <span>검수 체크리스트</span>
          <Button
            variant="primary"
            size="sm"
            onClick={handleSave}
            loading={saveInspections.isPending}
          >
            <Save size={14} />
            저장
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

        {/* 사진 URL */}
        <div className="flex items-center gap-3">
          <label className="text-sm text-neutral-500 w-20 shrink-0">사진 URL</label>
          <input
            type="text"
            value={current.photo_url}
            onChange={(e) => updateItem(activeIdx, 'photo_url', e.target.value)}
            placeholder="Google Drive 공유 링크..."
            className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm"
          />
        </div>
      </div>
    </Card>
  );
}
