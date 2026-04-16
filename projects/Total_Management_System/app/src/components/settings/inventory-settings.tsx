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

export default function InventorySettings({ settings, onSave, saving }: TabProps) {
  const [lowStock, setLowStock] = useState(3);
  const [imwebSync, setImwebSync] = useState(true);
  const [categories, setCategories] = useState<string[]>(['BL', 'TH', 'LO', 'SL', 'CB', 'CS', 'AC']);
  const [catLabels, setCatLabels] = useState<Record<string, string>>({ BL: '블런트', TH: '씨닝', LO: '롱', SL: '슬라이싱', CB: '빗', CS: '케이스', AC: '악세서리' });
  const [newCat, setNewCat] = useState('');
  const [newCatLabel, setNewCatLabel] = useState('');
  const [safetyStock, setSafetyStock] = useState(5);
  const [adjustReasons, setAdjustReasons] = useState<string[]>([]);
  const [newReason, setNewReason] = useState('');
  const [catTabVisible, setCatTabVisible] = useState<Record<string, boolean>>({});
  const [skuDigits, setSkuDigits] = useState(3);
  const [serialTrigger, setSerialTrigger] = useState('manual');
  const [stocktakeDay, setStocktakeDay] = useState(0);
  const [barcodeFormat, setBarcodeFormat] = useState('Code128');
  const [defaultSort, setDefaultSort] = useState('name');
  const [leadTimes, setLeadTimes] = useState<Record<string, number>>({});

  useEffect(() => {
    setLowStock(parse(settings['inventory.low_stock_threshold'], 3));
    setImwebSync(parse(settings['inventory.imweb_sync'], true));
    setCategories(parse(settings['inventory.categories'], ['BL', 'TH', 'LO', 'SL', 'CB', 'CS', 'AC']));
    setCatLabels(parse(settings['inventory.category_labels'], { BL: '블런트', TH: '씨닝', LO: '롱', SL: '슬라이싱', CB: '빗', CS: '케이스', AC: '악세서리' }));
    setCatTabVisible(parse(settings['inventory.category_tab_visible'], {}));
    setSafetyStock(parse(settings['inventory.safety_stock'], 5));
    setAdjustReasons(parse(settings['inventory.adjustment_reasons'], ['파손', '분실', '증정', '샘플', '실사 조정', '기타']));
    setSkuDigits(parse(settings['inventory.sku_digits'], 3));
    setSerialTrigger(parse(settings['inventory.serial_auto_trigger'], 'manual'));
    setStocktakeDay(parse(settings['inventory.stocktake_reminder_day'], 0));
    setBarcodeFormat(parse(settings['inventory.barcode_format'], 'Code128'));
    setDefaultSort(parse(settings['inventory.default_sort'], 'name'));
    setLeadTimes(parse(settings['inventory.lead_times'], {}));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'inventory.low_stock_threshold', value: lowStock },
      { key: 'dashboard.low_stock_threshold', value: lowStock },
      { key: 'inventory.imweb_sync', value: imwebSync },
      { key: 'inventory.categories', value: categories },
      { key: 'inventory.category_labels', value: catLabels },
      { key: 'inventory.category_tab_visible', value: catTabVisible },
      { key: 'inventory.safety_stock', value: safetyStock },
      { key: 'inventory.adjustment_reasons', value: adjustReasons },
      { key: 'inventory.sku_digits', value: skuDigits },
      { key: 'inventory.serial_auto_trigger', value: serialTrigger },
      { key: 'inventory.stocktake_reminder_day', value: stocktakeDay },
      { key: 'inventory.barcode_format', value: barcodeFormat },
      { key: 'inventory.default_sort', value: defaultSort },
      { key: 'inventory.lead_times', value: leadTimes },
    ]);
  };

  const CAT_LABELS: Record<string, string> = { BL: '블런트', TH: '씨닝', LO: '롱', SL: '슬라이싱', CB: '빗', CS: '케이스', AC: '악세서리' };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">상품·재고 설정</h2>

      <Field label="저재고 알림 기준 수량" desc="이 수량 이하 상품이 저재고로 표시됩니다.">
        <input type="number" value={lowStock} onChange={(e) => setLowStock(Number(e.target.value))}
          className="w-24 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
      </Field>

      <Field label="아임웹 재고 자동 동기화" desc="판매/취소 시 아임웹 재고도 연동합니다.">
        <Toggle checked={imwebSync} onChange={setImwebSync} />
      </Field>

      <Field label="상품 카테고리 목록" desc="코드는 수정 불가 (SKU 채번에 사용). 표시명만 변경 가능.">
        <div className="space-y-1.5 mb-3">
          {categories.map((code, i) => (
            <div key={i} className="flex items-center gap-2">
              <span className="w-10 text-xs font-mono font-bold text-neutral-700 bg-neutral-100 rounded px-2 py-1 text-center">{code}</span>
              <input
                value={catLabels[code] || ''}
                onChange={(e) => setCatLabels({ ...catLabels, [code]: e.target.value })}
                placeholder="표시명 입력"
                className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
              />
              <label className="flex items-center gap-1 text-[10px] text-neutral-400 shrink-0 cursor-pointer" title="제품 화면 탭에 표시">
                <input type="checkbox"
                  checked={catTabVisible[code] !== false}
                  onChange={(e) => setCatTabVisible({ ...catTabVisible, [code]: e.target.checked })}
                  className="rounded w-3.5 h-3.5" />
                탭
              </label>
              <button onClick={() => {
                setCategories(categories.filter((_, j) => j !== i));
                const next = { ...catLabels }; delete next[code]; setCatLabels(next);
              }} className="text-neutral-400 hover:text-red-500"><X size={14} /></button>
            </div>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newCat} onChange={(e) => setNewCat(e.target.value.toUpperCase())} placeholder="코드 (예: ST)"
            className="w-20 h-8 px-2 rounded-lg border border-neutral-200 text-sm font-mono" maxLength={3} />
          <input value={newCatLabel} onChange={(e) => setNewCatLabel(e.target.value)} placeholder="표시명 (예: 스트레이트)"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm" />
          <button onClick={() => {
            if (newCat.trim()) {
              setCategories([...categories, newCat.trim()]);
              if (newCatLabel.trim()) setCatLabels({ ...catLabels, [newCat.trim()]: newCatLabel.trim() });
              setNewCat(''); setNewCatLabel('');
            }
          }} className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      <Field label="안전재고 수량 (기본값)" desc="이 수량 이하면 '재주문 필요' 알림 표시.">
        <input type="number" value={safetyStock} onChange={(e) => setSafetyStock(Number(e.target.value))}
          className="w-24 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
      </Field>

      <Field label="재고 조정 사유 프리셋" desc="재고 조정 시 드롭다운에서 선택.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {adjustReasons.map((r, i) => (
            <span key={i} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
              {r}
              <button onClick={() => setAdjustReasons(adjustReasons.filter((_, j) => j !== i))} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
            </span>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newReason} onChange={(e) => setNewReason(e.target.value)} placeholder="새 사유"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            onKeyDown={(e) => { if (e.key === 'Enter' && newReason.trim()) { setAdjustReasons([...adjustReasons, newReason.trim()]); setNewReason(''); } }} />
          <button onClick={() => { if (newReason.trim()) { setAdjustReasons([...adjustReasons, newReason.trim()]); setNewReason(''); } }}
            className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      <Field label="SKU 자동 채번 자릿수" desc="카테고리 뒤 숫자 자릿수. 3이면 BL001.">
        <select value={skuDigits} onChange={(e) => setSkuDigits(Number(e.target.value))}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value={3}>3자리 (BL001)</option>
          <option value={4}>4자리 (BL0001)</option>
          <option value={5}>5자리 (BL00001)</option>
        </select>
      </Field>

      <Field label="시리얼 자동 생성 시점" desc="보관창고→준비창고 이동 시 자동 부여 권장.">
        <select value={serialTrigger} onChange={(e) => setSerialTrigger(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value="manual">수동 (직접 생성)</option>
          <option value="on_intake">입고 시 (보관창고)</option>
          <option value="on_move_to_ready">준비창고 이동 시</option>
        </select>
      </Field>

      <Field label="재고 실사 주기 알림" desc="매월 N일에 대시보드 리마인더. 0이면 비활성.">
        <div className="flex items-center gap-2">
          <span className="text-sm">매월</span>
          <input type="number" value={stocktakeDay} onChange={(e) => setStocktakeDay(Number(e.target.value))}
            className="w-16 h-9 px-2 rounded-lg border border-neutral-200 text-sm text-center" min={0} max={31} />
          <span className="text-sm">일</span>
        </div>
      </Field>

      <Field label="다중 창고 관리" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          보관창고 / 준비창고 분리 관리 기능이 추가됩니다. 재고 이동(보관→준비) 시 시리얼 자동 생성과 연동됩니다.
        </p>
      </Field>

      <Field label="바코드 형식" desc="">
        <select value={barcodeFormat} onChange={(e) => setBarcodeFormat(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value="Code128">Code128</option>
          <option value="QR">QR코드</option>
        </select>
      </Field>

      <Field label="상품 정렬 기본값" desc="제품 목록, 발주 작성 등에 적용됩니다.">
        <select value={defaultSort} onChange={(e) => setDefaultSort(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value="group">제품군 → 카테고리 → 순서 → 이름</option>
          <option value="name">이름순</option>
          <option value="category">카테고리순</option>
          <option value="stock_asc">재고 적은순</option>
          <option value="stock_desc">재고 많은순</option>
        </select>
      </Field>

      <Field label="매입처별 리드타임" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          매입처별 발주→입고 예상 일수. 매입처 등록 후 개별 설정 예정.
        </p>
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

function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <button onClick={() => onChange(!checked)}
      className={`relative w-11 h-6 rounded-full transition ${checked ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
      <span className={`absolute top-0.5 left-0.5 w-5 h-5 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-5' : ''}`} />
    </button>
  );
}
