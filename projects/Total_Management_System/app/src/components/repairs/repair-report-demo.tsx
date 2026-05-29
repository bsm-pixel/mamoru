'use client';

import { useState } from 'react';
import { Camera, Plus, Trash2 } from 'lucide-react';
import { InspectionMarkBoard } from './inspection-mark-board';
import { MarkOverlay } from './mark-overlay';
import { type MarkV2, colorOf, FLAG_TYPES } from './inspection-marks';
import { COMMON_MENT, COMMENT_PRESETS_BY_TYPE, TININ_INTRO } from '@/lib/repair/comment-presets';

const SCISSOR_TYPES = ['블런트', '틴닝', '장가위', '슬라이싱', '기타'];

/** 위치형 핀 라벨 → 자동 진단멘트 */
const PIN_MENT: Record<string, string> = {
  '무뎌짐': COMMON_MENT.dull,
  '찍힘': COMMON_MENT.nick,
  '빗살 손상': COMMON_MENT.comb,
  '부품 문제': COMMON_MENT.parts,
  '스토퍼 문제': COMMON_MENT.stopper,
};

interface DemoScissor {
  type: string;
  photoUrl: string;
  marks: MarkV2[];
  flags: string[];
  comment: string;
}

const newScissor = (): DemoScissor => ({ type: '블런트', photoUrl: '', marks: [], flags: [], comment: '' });

/* 멘트 블록(\n\n) 유틸 */
const splitBlocks = (c: string) => c.split('\n\n').map((b) => b.trim()).filter(Boolean);
const addBlocks = (c: string, blocks: string[]) => {
  const cur = splitBlocks(c);
  for (const b of blocks) { const t = b.trim(); if (t && !cur.includes(t)) cur.push(t); }
  return cur.join('\n\n');
};
const removeBlocks = (c: string, blocks: string[]) => {
  const rm = blocks.map((b) => b.trim());
  return splitBlocks(c).filter((b) => !rm.includes(b)).join('\n\n');
};
const flagKey = (note: string) => FLAG_TYPES.find((f) => f.note === note)?.key;
/** 플래그 ON 시 삽입할 멘트(틴닝+장력은 인트로 동반) */
const mentOn = (key: string | undefined, type: string): string[] => {
  if (key === 'tension') return type === '틴닝' ? [TININ_INTRO, COMMON_MENT.tension] : [COMMON_MENT.tension];
  if (key === 'balance') return [COMMON_MENT.balance];
  if (key === 'edgeangle') return [COMMON_MENT.edgeangle];
  return [];
};
const mentOff = (key: string | undefined): string[] => {
  if (key === 'tension') return [COMMON_MENT.tension];
  if (key === 'balance') return [COMMON_MENT.balance];
  if (key === 'edgeangle') return [COMMON_MENT.edgeangle];
  return [];
};

/**
 * 디자인모니터 데모 — 좌(검수 편집) / 우(실시간 고객 수리내역서).
 * 핀/플래그 선택 시 해당 진단멘트 자동삽입. 멘트는 가위별 독립, 우측 카드에 통합.
 */
export function RepairReportDemo() {
  const [scissors, setScissors] = useState<DemoScissor[]>([newScissor()]);
  const [activeIdx, setActiveIdx] = useState(0);

  const active = scissors[activeIdx];
  const presetList = COMMENT_PRESETS_BY_TYPE[active.type] || []; // 종류별 수동 칩(공통은 자동)

  const updateActive = (patch: Partial<DemoScissor>) =>
    setScissors((prev) => prev.map((s, i) => (i === activeIdx ? { ...s, ...patch } : s)));

  const addScissor = () => { setScissors((prev) => [...prev, newScissor()]); setActiveIdx(scissors.length); };
  const removeScissor = (idx: number) => {
    if (scissors.length <= 1) return;
    setScissors((prev) => prev.filter((_, i) => i !== idx));
    setActiveIdx((cur) => Math.max(0, cur - (idx <= cur ? 1 : 0)));
  };

  // 다시 촬영(새 사진) → 마킹·플래그·진단멘트 모두 초기화
  const pickPhoto = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => updateActive({ photoUrl: e.target?.result as string, marks: [], flags: [], comment: '' });
    reader.readAsDataURL(file);
  };

  // 핀 변경 → 라벨별(무뎌짐/찍힘/스토퍼) 멘트 자동삽입/해제
  const handleMarks = (newMarks: MarkV2[]) => {
    let c = active.comment;
    for (const [label, ment] of Object.entries(PIN_MENT)) {
      const had = active.marks.some((m) => m.label === label);
      const has = newMarks.some((m) => m.label === label);
      if (has && !had) c = addBlocks(c, [ment]);
      if (!has && had) c = removeBlocks(c, [ment]);
    }
    updateActive({ marks: newMarks, comment: c });
  };

  // 플래그 변경 → 해당 멘트 자동삽입/해제 (틴닝+장력은 인트로 동반)
  const handleFlags = (newFlags: string[]) => {
    const added = newFlags.filter((n) => !active.flags.includes(n));
    const removed = active.flags.filter((n) => !newFlags.includes(n));
    let c = active.comment;
    for (const n of added) c = addBlocks(c, mentOn(flagKey(n), active.type));
    for (const n of removed) c = removeBlocks(c, mentOff(flagKey(n)));
    updateActive({ flags: newFlags, comment: c });
  };

  // 수동 종류별 멘트 칩 토글
  const togglePreset = (p: string) => {
    const t = p.trim();
    const blocks = splitBlocks(active.comment);
    const i = blocks.indexOf(t);
    if (i >= 0) blocks.splice(i, 1); else blocks.push(t);
    updateActive({ comment: blocks.join('\n\n') });
  };
  const presetActive = (p: string) => splitBlocks(active.comment).includes(p.trim());

  const countsOf = (marks: MarkV2[]) => {
    const c: Record<string, number> = {};
    marks.forEach((m) => { c[m.label] = (c[m.label] || 0) + 1; });
    return c;
  };

  return (
    <div className="grid gap-5 lg:grid-cols-2">
      {/* ─── LEFT: 검수 편집 ─── */}
      <div>
        <div className="flex items-center gap-2 mb-2">
          <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-stone-800 text-white">사장님 입력 (검수)</span>
          <span className="text-[11px] text-stone-400">핀/플래그 선택 → 진단멘트 자동삽입</span>
        </div>
        <div className="mx-auto w-[390px] max-w-full bg-[#FAF9F7] rounded-[24px] shadow-lg border border-stone-200 overflow-hidden">
          <div className="p-4 space-y-4">
            {/* 가위 탭 */}
            <div className="flex gap-1 overflow-x-auto">
              {scissors.map((s, idx) => (
                <button key={idx} onClick={() => setActiveIdx(idx)}
                  className={`shrink-0 px-3 py-1.5 text-xs font-medium rounded-lg transition ${activeIdx === idx ? 'bg-neutral-800 text-white' : 'bg-neutral-100 text-neutral-600 hover:bg-neutral-200'}`}>
                  #{idx + 1} {s.type}
                </button>
              ))}
              <button onClick={addScissor} className="shrink-0 px-3 py-1.5 text-xs text-neutral-500 hover:text-neutral-800 flex items-center gap-1">
                <Plus size={12} /> 추가
              </button>
            </div>

            {/* 종류 + 삭제 */}
            <div className="flex items-center gap-3">
              <label className="text-sm text-neutral-500 w-12 shrink-0">종류</label>
              <select value={active.type} onChange={(e) => updateActive({ type: e.target.value })}
                className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm">
                {SCISSOR_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
              {scissors.length > 1 && (
                <button onClick={() => removeScissor(activeIdx)} className="text-red-400 hover:text-red-600"><Trash2 size={14} /></button>
              )}
            </div>

            {/* 사진 */}
            {!active.photoUrl ? (
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

            {/* 핀 마킹 (controlled, 자동 멘트 연동) */}
            <div>
              <label className="block text-sm text-neutral-500 mb-2">상처 표시</label>
              <InspectionMarkBoard photoUrl={active.photoUrl} marks={active.marks} flags={active.flags} onMarks={handleMarks} onFlags={handleFlags} scissorType={active.type} />
            </div>

            {/* 진단 멘트 — 종류별 수동 칩 (공통은 핀/플래그로 자동 삽입) */}
            <div className="pt-3 border-t border-neutral-100">
              <label className="block text-sm font-medium text-neutral-600 mb-2">진단 멘트 — #{activeIdx + 1} {active.type}</label>
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
              <p className="text-[11px] text-stone-400 mb-1.5">※ 장력조절·스토퍼·밸런스·날각은 위에서 표시하면 멘트가 자동으로 들어갑니다.</p>
              <textarea value={active.comment} onChange={(e) => updateActive({ comment: e.target.value })} rows={6}
                placeholder="칩 선택·자동삽입된 멘트가 여기 모입니다. 직접 수정도 가능. 줄바꿈 그대로 고객에게 표시됩니다."
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-white text-sm leading-relaxed resize-y focus:outline-none focus:border-neutral-400" />
            </div>
          </div>
        </div>
      </div>

      {/* ─── RIGHT: 고객 수리내역서 (실시간) ─── */}
      <div>
        <div className="flex items-center gap-2 mb-2">
          <span className="text-xs font-semibold px-2.5 py-1 rounded-full bg-emerald-600 text-white">고객 화면 (실시간)</span>
          <span className="text-[11px] text-stone-400">알림톡 수리내역서로 보이는 형태</span>
        </div>
        <div className="mx-auto w-[390px] max-w-full bg-[#FAF9F7] rounded-[24px] shadow-lg border border-stone-200 overflow-hidden">
          <div>
            <div className="bg-[#1A1A1A] text-[#FAF9F7] px-4 py-4 text-center">
              <div className="text-[11px] tracking-widest opacity-70">MAMORU</div>
              <div className="text-base font-bold mt-0.5">수리내역서</div>
            </div>

            {scissors.map((s, i) => {
              const counts = countsOf(s.marks);
              return (
                <div key={i} className="m-4 bg-white rounded-2xl border border-[#D4D0CB] overflow-hidden shadow-sm">
                  <div className="bg-[#1A1A1A] text-[#FAF9F7] px-4 py-3 flex items-center justify-between">
                    <span className="font-bold">{i + 1}번 가위</span>
                    <span className="text-[12px] px-2.5 py-0.5 rounded-full bg-white/15">{s.type}</span>
                  </div>
                  <div className="p-4 text-center">
                    <div className="inline-block w-full max-w-[280px]">
                      <MarkOverlay photoUrl={s.photoUrl} marks={s.marks} flags={s.flags} />
                    </div>
                    {Object.keys(counts).length > 0 && (
                      <div className="mt-3 flex flex-wrap gap-1.5 justify-center">
                        {Object.entries(counts).map(([label, n]) => (
                          <span key={label} className="flex items-center gap-1.5 text-[12px] text-[#4A4A4A] bg-white border border-[#D4D0CB] rounded-full px-2.5 py-0.5">
                            <span className="w-2.5 h-2.5 rounded-full" style={{ background: colorOf(label) }} />{label} {n}
                          </span>
                        ))}
                      </div>
                    )}
                    {s.marks.length === 0 && s.flags.length === 0 && s.photoUrl && <div className="mt-2 text-[13px] text-[#8A8580]">표시된 상처 없음 (양호)</div>}
                  </div>
                  {s.comment.trim() && (
                    <div className="px-4 pb-4 -mt-1">
                      <div className="rounded-xl bg-[#F5F3F0] border border-[#E7E2DC] px-3.5 py-3">
                        <div className="text-[11px] font-bold tracking-wide text-[#8A8580] mb-2">진단 및 내역</div>
                        <div className="space-y-2.5">
                          {s.comment.split('\n\n').map((b) => b.trim()).filter(Boolean).map((b, bi) => (
                            <div key={bi} className="flex gap-2 text-[13.5px] text-[#3D3A36] leading-[1.7]">
                              <span className="text-[#C2B6A6] shrink-0 select-none mt-px">–</span>
                              <span className="whitespace-pre-wrap">{b}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
            <div className="h-2" />
          </div>
        </div>
      </div>
    </div>
  );
}
