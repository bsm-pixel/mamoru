'use client';

import { useState } from 'react';
import { Camera } from 'lucide-react';
import { InspectionMarkBoard } from './inspection-mark-board';
import { MarkOverlay } from './mark-overlay';
import { type MarkV2, colorOf } from './inspection-marks';
import { COMMENT_PRESETS_COMMON, COMMENT_PRESETS_BY_TYPE } from '@/lib/repair/comment-presets';

const SCISSOR_TYPES = ['블런트', '틴닝', '장가위', '슬라이싱', '기타'];

/**
 * 디자인모니터 데모 — 좌(검수 편집) / 우(실시간 고객 수리내역서 미리보기).
 * 좌측에서 마킹/멘트를 바꾸면 우측 고객 화면이 즉시 갱신된다. (저장/업로드 없음)
 */
export function RepairReportDemo() {
  const [photoUrl, setPhotoUrl] = useState('');
  const [scissorType, setScissorType] = useState('블런트');
  const [marks, setMarks] = useState<MarkV2[]>([]);
  const [flags, setFlags] = useState<string[]>([]);
  const [comment, setComment] = useState('');

  const presetList = [...COMMENT_PRESETS_COMMON, ...(COMMENT_PRESETS_BY_TYPE[scissorType] || [])];

  const togglePreset = (p: string) => {
    const t = p.trim();
    setComment((prev) => {
      const blocks = prev.split('\n\n').map((b) => b.trim()).filter(Boolean);
      const i = blocks.indexOf(t);
      if (i >= 0) blocks.splice(i, 1); else blocks.push(t);
      return blocks.join('\n\n');
    });
  };
  const presetActive = (p: string) => comment.split('\n\n').map((b) => b.trim()).includes(p.trim());

  const pickPhoto = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      setPhotoUrl(e.target?.result as string);
      setMarks([]); setFlags([]); // 새 사진 → 마킹 초기화
    };
    reader.readAsDataURL(file);
  };

  const counts: Record<string, number> = {};
  marks.forEach((m) => { counts[m.label] = (counts[m.label] || 0) + 1; });

  return (
    <div className="grid gap-5 lg:grid-cols-2">
      {/* ─── LEFT: 검수 편집 (모바일) ─── */}
      <div>
        <div className="flex items-center gap-2 mb-2">
          <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-stone-800 text-white">사장님 입력 (검수)</span>
          <span className="text-[11px] text-stone-400">체크/마킹하면 → 우측 고객 화면 실시간 반영</span>
        </div>
        <div className="mx-auto w-[390px] max-w-full bg-[#FAF9F7] rounded-[24px] shadow-lg border border-stone-200 overflow-hidden">
          <div className="p-4 space-y-4 max-h-[76vh] overflow-y-auto">
            {/* 종류 */}
            <div className="flex items-center gap-3">
              <label className="text-sm text-neutral-500 w-12 shrink-0">종류</label>
              <select value={scissorType} onChange={(e) => setScissorType(e.target.value)}
                className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm">
                {SCISSOR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            </div>

            {/* 사진 */}
            {!photoUrl ? (
              <label className="cursor-pointer flex flex-col items-center justify-center w-28 h-28 border-2 border-dashed border-neutral-300 rounded-lg hover:border-neutral-400 transition">
                <Camera size={24} className="text-neutral-400 mb-1" />
                <span className="text-[11px] text-neutral-400">촬영/업로드</span>
                <input type="file" accept="image/*" capture="environment" className="hidden"
                  onChange={(e) => { const f = e.target.files?.[0]; if (f) pickPhoto(f); e.target.value = ''; }} />
              </label>
            ) : (
              <div className="flex items-center gap-2">
                <span className="text-[11px] text-neutral-400">사진 변경:</span>
                <label className="cursor-pointer inline-flex items-center gap-1 px-2.5 py-1 rounded-lg border border-neutral-200 text-xs text-neutral-600 hover:border-neutral-400">
                  <Camera size={13} /> 다시 촬영
                  <input type="file" accept="image/*" capture="environment" className="hidden"
                    onChange={(e) => { const f = e.target.files?.[0]; if (f) pickPhoto(f); e.target.value = ''; }} />
                </label>
              </div>
            )}

            {/* 핀 마킹 (controlled) */}
            <div>
              <label className="block text-sm text-neutral-500 mb-2">상처 표시</label>
              <InspectionMarkBoard photoUrl={photoUrl} marks={marks} flags={flags} onMarks={setMarks} onFlags={setFlags} />
            </div>

            {/* 진단 멘트 */}
            <div className="pt-3 border-t border-neutral-100">
              <label className="block text-sm font-medium text-neutral-600 mb-2">진단 멘트 ({scissorType})</label>
              <div className="flex gap-1.5 flex-wrap mb-2">
                {presetList.map((p) => {
                  const active = presetActive(p);
                  return (
                    <button key={p} type="button" onClick={() => togglePreset(p)} title={p}
                      className={`px-2.5 py-1 text-[11px] rounded-lg border text-left leading-snug transition max-w-[300px] truncate ${active ? 'bg-neutral-800 text-white border-neutral-800' : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-300'}`}>
                      {p.split('\n')[0]}{p.includes('\n') ? ' …' : ''}
                    </button>
                  );
                })}
              </div>
              <textarea value={comment} onChange={(e) => setComment(e.target.value)} rows={5}
                placeholder="멘트 칩을 선택하거나 직접 작성. 줄바꿈 그대로 고객에게 표시됩니다."
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-white text-sm leading-relaxed resize-y focus:outline-none focus:border-neutral-400" />
            </div>
          </div>
        </div>
      </div>

      {/* ─── RIGHT: 고객이 보는 수리내역서 (실시간) ─── */}
      <div>
        <div className="flex items-center gap-2 mb-2">
          <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-emerald-600 text-white">고객 화면 (실시간)</span>
          <span className="text-[11px] text-stone-400">알림톡 수리내역서로 보이는 형태</span>
        </div>
        <div className="mx-auto w-[390px] max-w-full bg-[#FAF9F7] rounded-[24px] shadow-lg border border-stone-200 overflow-hidden">
          <div className="max-h-[76vh] overflow-y-auto">
            {/* 헤더 */}
            <div className="bg-[#1A1A1A] text-[#FAF9F7] px-4 py-4 text-center">
              <div className="text-[11px] tracking-widest opacity-70">MAMORU</div>
              <div className="text-base font-bold mt-0.5">수리내역서</div>
            </div>

            {/* 가위 카드 */}
            <div className="m-4 bg-white rounded-2xl border border-[#D4D0CB] overflow-hidden shadow-sm">
              <div className="bg-[#1A1A1A] text-[#FAF9F7] px-4 py-3 flex items-center justify-between">
                <span className="font-bold">1번 가위</span>
                <span className="text-[12px] px-2.5 py-0.5 rounded-full bg-white/15">{scissorType}</span>
              </div>
              <div className="p-4 text-center">
                <div className="inline-block w-full max-w-[280px]">
                  <MarkOverlay photoUrl={photoUrl} marks={marks} flags={flags} />
                </div>
                {/* 범례 */}
                {Object.keys(counts).length > 0 && (
                  <div className="mt-3 flex flex-wrap gap-1.5 justify-center">
                    {Object.entries(counts).map(([label, n]) => (
                      <span key={label} className="flex items-center gap-1.5 text-[12px] text-[#4A4A4A] bg-white border border-[#D4D0CB] rounded-full px-2.5 py-0.5">
                        <span className="w-2.5 h-2.5 rounded-full" style={{ background: colorOf(label) }} />{label} {n}
                      </span>
                    ))}
                  </div>
                )}
                {marks.length === 0 && photoUrl && <div className="mt-2 text-[13px] text-[#8A8580]">표시된 상처 없음 (양호)</div>}
              </div>
            </div>

            {/* 진단 멘트 */}
            {comment.trim() ? (
              <div className="mx-4 mb-5 p-4 bg-[#FAF9F7] rounded-xl border-l-4 border-[#1A1A1A]">
                <div className="text-[13px] font-bold text-[#1A1A1A] mb-2 pb-2 border-b border-[#D4D0CB]">진단 안내</div>
                <div className="text-[14px] text-[#4A4A4A] leading-[1.7] whitespace-pre-wrap">{comment}</div>
              </div>
            ) : (
              <div className="mx-4 mb-5 p-4 text-center text-[13px] text-[#8A8580]">진단 멘트를 입력하면 여기에 표시됩니다</div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
