'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Plus, X } from 'lucide-react';
import type { TabProps } from '@/app/(dashboard)/settings/page';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

interface B2BCategory {
  key: string;
  label: string;
  icon: string;
  display_order: number;
  is_active: boolean;
  is_default?: boolean;
}

const ICON_OPTIONS: { value: string; label: string }[] = [
  { value: 'Users',         label: '👥 사람들 (딜러/도매)' },
  { value: 'GraduationCap', label: '🎓 졸업모 (아카데미/교육)' },
  { value: 'School',        label: '🏫 학교' },
  { value: 'Building2',     label: '🏢 건물 (회사/매입처)' },
  { value: 'Building',      label: '🏛 관공서/공기관' },
  { value: 'Briefcase',     label: '💼 사업자' },
  { value: 'Hospital',      label: '🏥 병원' },
  { value: 'Store',         label: '🏪 매장' },
];

const DEFAULT_B2B_CATEGORIES: B2BCategory[] = [
  { key: 'dealer',  label: '딜러',     icon: 'Users',          display_order: 1, is_active: true, is_default: true },
  { key: 'academy', label: '아카데미', icon: 'GraduationCap',  display_order: 2, is_active: true, is_default: true },
];

export default function CustomerSettings({ settings, onSave, saving }: TabProps) {
  const [types, setTypes] = useState<string[]>(['retail', 'online', 'dealer', 'academy']);
  const [newType, setNewType] = useState('');
  const [sources, setSources] = useState<string[]>(['imweb', 'consultation', 'as', 'manual']);
  const [newSource, setNewSource] = useState('');
  const [rfm, setRfm] = useState({ recency_vip: 90, recency_dormant: 180, frequency: 3, monetary: 1000000 });
  const [outstandingDays, setOutstandingDays] = useState(0);
  const [defaultSort, setDefaultSort] = useState('name');
  const [memoTemplates, setMemoTemplates] = useState<string[]>([]);
  const [newMemo, setNewMemo] = useState('');
  const [tags, setTags] = useState<string[]>([]);
  const [newTag, setNewTag] = useState('');
  // 074: B2B 카테고리 동적 관리
  const [b2bCategories, setB2bCategories] = useState<B2BCategory[]>(DEFAULT_B2B_CATEGORIES);

  useEffect(() => {
    setTypes(parse(settings['customer.types'], ['retail', 'online', 'dealer', 'academy']));
    setSources(parse(settings['customer.sources'], ['imweb', 'consultation', 'as', 'manual']));
    setRfm(parse(settings['customer.rfm'], { recency_vip: 90, recency_dormant: 180, frequency: 3, monetary: 1000000 }));
    setOutstandingDays(parse(settings['customer.outstanding_reminder_days'], 0));
    setDefaultSort(parse(settings['customer.default_sort'], 'name'));
    setMemoTemplates(parse(settings['customer.memo_templates'], []));
    setTags(parse(settings['customer.tags'], []));
    setB2bCategories(parse(settings['b2b.categories'], DEFAULT_B2B_CATEGORIES));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'customer.types', value: types },
      { key: 'customer.sources', value: sources },
      { key: 'customer.rfm', value: rfm },
      { key: 'customer.outstanding_reminder_days', value: outstandingDays },
      { key: 'customer.default_sort', value: defaultSort },
      { key: 'customer.memo_templates', value: memoTemplates },
      { key: 'customer.tags', value: tags },
      { key: 'b2b.categories', value: b2bCategories },
    ]);
  };

  function addB2BCategory() {
    const nextOrder = Math.max(0, ...b2bCategories.map((c) => c.display_order)) + 1;
    setB2bCategories([
      ...b2bCategories,
      { key: '', label: '', icon: 'Briefcase', display_order: nextOrder, is_active: true, is_default: false },
    ]);
  }
  function updateB2BCategory(idx: number, patch: Partial<B2BCategory>) {
    setB2bCategories(b2bCategories.map((c, i) => i === idx ? { ...c, ...patch } : c));
  }
  function removeB2BCategory(idx: number) {
    if (b2bCategories[idx]?.is_default) return; // is_default는 삭제 차단
    setB2bCategories(b2bCategories.filter((_, i) => i !== idx));
  }

  const TYPE_LABELS: Record<string, string> = { retail: '일반', online: '온라인', dealer: '딜러', academy: '아카데미' };
  const SOURCE_LABELS: Record<string, string> = { imweb: '아임웹', consultation: '상담', as: '복원수리', manual: '수동' };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">고객 관리 설정</h2>

      {/* 1. 고객 유형 */}
      <Field label="고객 유형 목록" desc="고객 등록/검색 시 드롭다운에 표시됩니다.">
        <ChipList items={types} labels={TYPE_LABELS} onRemove={(i) => setTypes(types.filter((_, j) => j !== i))} />
        <AddInput value={newType} onChange={setNewType} onAdd={() => { if (newType.trim()) { setTypes([...types, newType.trim()]); setNewType(''); } }} placeholder="새 유형" />
      </Field>

      {/* 074: B2B 카테고리 관리 */}
      <Field label="B2B 납품처 카테고리" desc="딜러/아카데미/학교/공기관 등 B2B 거래처 분류. /거래처 페이지 탭에 자동 노출. 신규 카테고리 추가 시 '고객 유형 목록'에도 같은 key 추가 + 단가 그룹(판매 관리 설정)에도 등록 권장.">
        <div className="space-y-2">
          {b2bCategories.sort((a, b) => a.display_order - b.display_order).map((cat, idx) => (
            <div key={idx} className="grid grid-cols-[80px_1fr_1fr_1fr_60px_60px_36px] gap-2 items-center p-2 rounded-lg border border-neutral-200 bg-neutral-50">
              <input
                type="text" value={cat.key}
                onChange={(e) => updateB2BCategory(b2bCategories.indexOf(cat), { key: e.target.value.replace(/[^a-z0-9_]/gi, '').toLowerCase() })}
                placeholder="key (영문)"
                disabled={cat.is_default}
                className="h-8 px-2 rounded border border-neutral-200 text-xs disabled:bg-neutral-100 disabled:text-neutral-400"
              />
              <input
                type="text" value={cat.label}
                onChange={(e) => updateB2BCategory(b2bCategories.indexOf(cat), { label: e.target.value })}
                placeholder="표시명 (예: 학교)"
                className="h-8 px-2 rounded border border-neutral-200 text-xs"
              />
              <select
                value={cat.icon}
                onChange={(e) => updateB2BCategory(b2bCategories.indexOf(cat), { icon: e.target.value })}
                className="h-8 px-1 rounded border border-neutral-200 text-xs"
              >
                {ICON_OPTIONS.map((opt) => (
                  <option key={opt.value} value={opt.value}>{opt.label}</option>
                ))}
              </select>
              <input
                type="number" value={cat.display_order}
                onChange={(e) => updateB2BCategory(b2bCategories.indexOf(cat), { display_order: Number(e.target.value) })}
                className="h-8 px-2 rounded border border-neutral-200 text-xs"
                title="순서"
              />
              <label className="flex items-center justify-center gap-1 text-[10px] text-neutral-500">
                <input
                  type="checkbox" checked={cat.is_active}
                  onChange={(e) => updateB2BCategory(b2bCategories.indexOf(cat), { is_active: e.target.checked })}
                />
                활성
              </label>
              <span className="text-[10px] text-neutral-400 text-center">
                {cat.is_default ? '기본' : ''}
              </span>
              <button
                onClick={() => removeB2BCategory(b2bCategories.indexOf(cat))}
                disabled={cat.is_default}
                className="text-neutral-400 hover:text-red-500 disabled:text-neutral-200 disabled:cursor-not-allowed"
                title={cat.is_default ? '기본 카테고리는 삭제 불가 (비활성 토글로 숨김 가능)' : '삭제'}
              >
                <X size={14} />
              </button>
            </div>
          ))}
          <button onClick={addB2BCategory} className="flex items-center gap-1 px-3 py-1.5 rounded-lg border border-dashed border-neutral-300 text-xs text-neutral-500 hover:bg-neutral-50">
            <Plus size={12} />새 카테고리 추가
          </button>
        </div>
      </Field>

      {/* 2. RFM 분류 기준 */}
      <Field label="RFM 분류 기준" desc="VIP/일반/휴면 고객 세그먼트를 결정하는 기준값.">
        <div className="grid grid-cols-2 gap-3">
          <div>
            <span className="text-xs text-neutral-500">VIP 기준: 최근 거래</span>
            <div className="flex items-center gap-1 mt-1">
              <input type="number" value={rfm.recency_vip} onChange={(e) => setRfm({ ...rfm, recency_vip: Number(e.target.value) })}
                className="w-16 h-8 px-2 rounded-lg border border-neutral-200 text-sm text-center" />
              <span className="text-xs">일 이내</span>
            </div>
          </div>
          <div>
            <span className="text-xs text-neutral-500">휴면 기준: 최근 거래</span>
            <div className="flex items-center gap-1 mt-1">
              <input type="number" value={rfm.recency_dormant} onChange={(e) => setRfm({ ...rfm, recency_dormant: Number(e.target.value) })}
                className="w-16 h-8 px-2 rounded-lg border border-neutral-200 text-sm text-center" />
              <span className="text-xs">일 초과</span>
            </div>
          </div>
          <div>
            <span className="text-xs text-neutral-500">빈도 기준</span>
            <div className="flex items-center gap-1 mt-1">
              <input type="number" value={rfm.frequency} onChange={(e) => setRfm({ ...rfm, frequency: Number(e.target.value) })}
                className="w-16 h-8 px-2 rounded-lg border border-neutral-200 text-sm text-center" />
              <span className="text-xs">회 이상</span>
            </div>
          </div>
          <div>
            <span className="text-xs text-neutral-500">금액 기준</span>
            <div className="flex items-center gap-1 mt-1">
              <input type="number" value={rfm.monetary} onChange={(e) => setRfm({ ...rfm, monetary: Number(e.target.value) })}
                className="w-24 h-8 px-2 rounded-lg border border-neutral-200 text-sm" step={100000} />
              <span className="text-xs">원 이상</span>
            </div>
          </div>
        </div>
      </Field>

      {/* 3. 미수금 독촉 */}
      <Field label="미수금 독촉 자동 알림" desc="N일 경과 시 대시보드 알림. 0이면 비활성.">
        <div className="flex items-center gap-2">
          <input type="number" value={outstandingDays} onChange={(e) => setOutstandingDays(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      {/* 7. 고객 소스 */}
      <Field label="고객 소스 (유입 경로)" desc="고객 등록 시 선택 가능한 유입 경로.">
        <ChipList items={sources} labels={SOURCE_LABELS} onRemove={(i) => setSources(sources.filter((_, j) => j !== i))} />
        <AddInput value={newSource} onChange={setNewSource} onAdd={() => { if (newSource.trim()) { setSources([...sources, newSource.trim()]); setNewSource(''); } }} placeholder="새 소스" />
      </Field>

      {/* 8. 기본 정렬 */}
      <Field label="고객 목록 기본 정렬" desc="">
        <select value={defaultSort} onChange={(e) => setDefaultSort(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value="name">이름순</option>
          <option value="recent">최근 거래순</option>
          <option value="outstanding">미수금순</option>
        </select>
      </Field>

      {/* 10. 메모 템플릿 */}
      <Field label="고객 메모 템플릿" desc="자주 쓰는 메모를 미리 등록해두면 원클릭 삽입 가능.">
        <div className="space-y-1 mb-2">
          {memoTemplates.map((m, i) => (
            <div key={i} className="flex items-center gap-2">
              <span className="flex-1 text-sm bg-neutral-50 rounded px-2 py-1">{m}</span>
              <button onClick={() => setMemoTemplates(memoTemplates.filter((_, j) => j !== i))} className="text-neutral-400 hover:text-red-500"><X size={14} /></button>
            </div>
          ))}
        </div>
        <AddInput value={newMemo} onChange={setNewMemo} onAdd={() => { if (newMemo.trim()) { setMemoTemplates([...memoTemplates, newMemo.trim()]); setNewMemo(''); } }} placeholder="새 메모 템플릿" />
      </Field>

      {/* 18. 태그 */}
      <Field label="고객 태그/라벨" desc="고객 프로파일에 부착할 수 있는 태그.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {tags.map((t, i) => (
            <span key={i} className="flex items-center gap-1 px-2 py-0.5 text-xs bg-blue-50 text-blue-700 rounded-full">
              {t}
              <button onClick={() => setTags(tags.filter((_, j) => j !== i))} className="hover:text-red-500"><X size={10} /></button>
            </span>
          ))}
        </div>
        <AddInput value={newTag} onChange={setNewTag} onAdd={() => { if (newTag.trim()) { setTags([...tags, newTag.trim()]); setNewTag(''); } }} placeholder="새 태그" />
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}><Save size={14} />{saving ? '저장 중...' : '저장'}</Button>
      </div>
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (<div><label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>{desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}{children}</div>);
}

function ChipList({ items, labels, onRemove }: { items: string[]; labels: Record<string, string>; onRemove: (i: number) => void }) {
  return (
    <div className="flex flex-wrap gap-1.5 mb-2">
      {items.map((item, i) => (
        <span key={i} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
          {labels[item] || item}
          <button onClick={() => onRemove(i)} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
        </span>
      ))}
    </div>
  );
}

function AddInput({ value, onChange, onAdd, placeholder }: { value: string; onChange: (v: string) => void; onAdd: () => void; placeholder: string }) {
  return (
    <div className="flex gap-2">
      <input value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder}
        className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
        onKeyDown={(e) => { if (e.key === 'Enter') onAdd(); }} />
      <button onClick={onAdd} className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
    </div>
  );
}
