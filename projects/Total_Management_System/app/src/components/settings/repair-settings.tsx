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

interface ShippingFee { qty: number; fee: number; }
interface ExtraService { name: string; price: number; }
interface BankAccount { bank: string; number: string; holder: string; }

export default function RepairSettings({ settings, onSave, saving }: TabProps) {
  const [priceMamoru, setPriceMamoru] = useState(10000);
  const [priceOther, setPriceOther] = useState(20000);
  const [shippingFees, setShippingFees] = useState<ShippingFee[]>([
    { qty: 1, fee: 5000 }, { qty: 2, fee: 3000 }, { qty: 3, fee: 0 },
  ]);
  const [extraServices, setExtraServices] = useState<ExtraService[]>([]);
  const [newServiceName, setNewServiceName] = useState('');
  const [newServicePrice, setNewServicePrice] = useState(0);
  const [bankAccount, setBankAccount] = useState<BankAccount>({ bank: '', number: '', holder: '' });
  const [staleDays, setStaleDays] = useState(3);
  const [inspectionCategories, setInspectionCategories] = useState<Record<string, string[]>>({});
  const [unpaidReminderDays, setUnpaidReminderDays] = useState(3);
  const [estimatedDays, setEstimatedDays] = useState(7);

  useEffect(() => {
    setPriceMamoru(parse(settings['repair.price_mamoru'], 10000));
    setPriceOther(parse(settings['repair.price_other'], 20000));
    setShippingFees(parse(settings['repair.shipping_fees'], [
      { qty: 1, fee: 5000 }, { qty: 2, fee: 3000 }, { qty: 3, fee: 0 },
    ]));
    setExtraServices(parse(settings['repair.extra_services'], [{ name: '날 변형 (곡률 조절)', price: 50000 }]));
    setBankAccount(parse(settings['repair.bank_account'], { bank: '', number: '', holder: '' }));
    setStaleDays(parse(settings['repair.stale_days'], 3));
    setInspectionCategories(parse(settings['repair.inspection_categories'], {}));
    setUnpaidReminderDays(parse(settings['repair.unpaid_reminder_days'], 3));
    setEstimatedDays(parse(settings['repair.estimated_days'], 7));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'repair.price_mamoru', value: priceMamoru },
      { key: 'repair.price_other', value: priceOther },
      { key: 'repair.shipping_fees', value: shippingFees },
      { key: 'repair.extra_services', value: extraServices },
      { key: 'repair.bank_account', value: bankAccount },
      { key: 'repair.stale_days', value: staleDays },
      { key: 'repair.inspection_categories', value: inspectionCategories },
      { key: 'repair.unpaid_reminder_days', value: unpaidReminderDays },
      { key: 'repair.estimated_days', value: estimatedDays },
    ]);
  };

  const addExtraService = () => {
    if (!newServiceName.trim()) return;
    setExtraServices([...extraServices, { name: newServiceName.trim(), price: newServicePrice }]);
    setNewServiceName(''); setNewServicePrice(0);
  };

  const INSPECTION_LABELS: Record<string, string> = {
    blade_tip: '날 끝부', blade_mid: '날 중간', blade_inner: '날 안쪽',
    comb: '빗살', tension: '장력', parts: '내부부품', stopper: '스토퍼',
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">복원수리 설정</h2>

      {/* 1. 마모루 수리비 */}
      <Field label="마모루 가위 수리비 (개당)" desc="비용안내 자동 계산에 적용됩니다.">
        <div className="flex items-center gap-2">
          <input type="number" value={priceMamoru} onChange={(e) => setPriceMamoru(Number(e.target.value))}
            className="w-32 h-9 px-3 rounded-lg border border-neutral-200 text-sm" step={1000} />
          <span className="text-sm text-neutral-500">원</span>
        </div>
      </Field>

      {/* 2. 타사 수리비 */}
      <Field label="타사 가위 수리비 (개당)" desc="">
        <div className="flex items-center gap-2">
          <input type="number" value={priceOther} onChange={(e) => setPriceOther(Number(e.target.value))}
            className="w-32 h-9 px-3 rounded-lg border border-neutral-200 text-sm" step={1000} />
          <span className="text-sm text-neutral-500">원</span>
        </div>
      </Field>

      {/* 3. 수거 배송비 */}
      <Field label="수거 배송비 (수량별)" desc="방문수거 시 적용. 수량이 클수록 저렴.">
        <div className="space-y-1.5">
          {shippingFees.map((sf, i) => (
            <div key={i} className="flex items-center gap-2">
              <span className="text-sm w-16">{sf.qty}개{sf.qty >= 3 ? '+' : ''}</span>
              <input type="number" value={sf.fee} onChange={(e) => {
                const next = [...shippingFees]; next[i] = { ...sf, fee: Number(e.target.value) }; setShippingFees(next);
              }} className="w-24 h-8 px-2 rounded-lg border border-neutral-200 text-sm" step={1000} />
              <span className="text-xs text-neutral-500">원</span>
            </div>
          ))}
        </div>
      </Field>

      {/* 4. 추가 서비스 목록 (유동) */}
      <Field label="추가 서비스 목록" desc="날 변형, 슬라이싱 가공 등. 비용안내 시 선택 가능.">
        <div className="space-y-1.5 mb-2">
          {extraServices.map((s, i) => (
            <div key={i} className="flex items-center gap-2">
              <span className="flex-1 text-sm">{s.name}</span>
              <span className="text-sm font-mono">{s.price.toLocaleString()}원</span>
              <button onClick={() => setExtraServices(extraServices.filter((_, j) => j !== i))}
                className="text-neutral-400 hover:text-red-500"><X size={14} /></button>
            </div>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newServiceName} onChange={(e) => setNewServiceName(e.target.value)} placeholder="서비스명"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input type="number" value={newServicePrice || ''} onChange={(e) => setNewServicePrice(Number(e.target.value))} placeholder="금액"
            className="w-24 h-8 px-2 rounded-lg border border-neutral-200 text-sm" step={1000} />
          <button onClick={addExtraService} className="px-2 py-1 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 5. 입금 계좌 */}
      <Field label="입금 계좌 정보" desc="비용안내 알림톡/페이지에 표시됩니다.">
        <div className="grid grid-cols-3 gap-2">
          <input placeholder="은행명" value={bankAccount.bank} onChange={(e) => setBankAccount({ ...bankAccount, bank: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="계좌번호" value={bankAccount.number} onChange={(e) => setBankAccount({ ...bankAccount, number: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="예금주" value={bankAccount.holder} onChange={(e) => setBankAccount({ ...bankAccount, holder: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>
      </Field>

      {/* 6. 체류 경고 일수 */}
      <Field label="체류 경고 일수" desc="상태 변경 없이 이 일수 넘으면 '정체' 표시.">
        <div className="flex items-center gap-2">
          <input type="number" value={staleDays} onChange={(e) => setStaleDays(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={1} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      {/* 8. 검수 항목 카테고리 */}
      <Field label="검수 항목 카테고리" desc="각 부위별 선택 가능한 증상 목록. 다중 선택 가능.">
        {Object.entries(inspectionCategories).length === 0 ? (
          <p className="text-xs text-neutral-400">설정 데이터 없음 (DB 마이그레이션 후 표시)</p>
        ) : (
          <div className="space-y-3">
            {Object.entries(inspectionCategories).map(([part, options]) => (
              <div key={part}>
                <span className="text-sm font-medium">{INSPECTION_LABELS[part] || part}</span>
                <div className="flex flex-wrap gap-1 mt-1">
                  {(options as string[]).map((opt, i) => (
                    <span key={i} className="px-2 py-0.5 text-xs bg-neutral-100 rounded">{opt}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
        <p className="text-xs text-neutral-400 mt-2">검수 항목 추가/삭제 기능은 다음 업데이트에서 제공됩니다.</p>
      </Field>

      {/* 15. 미입금 리마인더 */}
      <Field label="미입금 리마인더" desc="비용안내 후 N일 미입금 시 자동 리마인더. 0이면 비활성.">
        <div className="flex items-center gap-2">
          <input type="number" value={unpaidReminderDays} onChange={(e) => setUnpaidReminderDays(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      {/* 20. 수리 완료 예상 기간 */}
      <Field label="수리 완료 예상 기간" desc="접수 확인 시 고객에게 안내되는 기본 소요 기간.">
        <div className="flex items-center gap-2">
          <input type="number" value={estimatedDays} onChange={(e) => setEstimatedDays(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={1} />
          <span className="text-sm text-neutral-500">일</span>
        </div>
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}>
          <Save size={14} />
          {saving ? '저장 중...' : '저장'}
        </Button>
      </div>
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (
    <div>
      <label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>
      {desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}
      {children}
    </div>
  );
}
