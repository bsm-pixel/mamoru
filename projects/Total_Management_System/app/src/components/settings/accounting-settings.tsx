'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Plus, X } from 'lucide-react';
import type { TabProps } from '@/app/(dashboard)/settings/page';
import {
  DEFAULT_EXPENSE_CATEGORIES,
  DEFAULT_CASHFLOW_INCOME_CATEGORIES,
  DEFAULT_CASHFLOW_EXPENSE_CATEGORIES,
} from '@/lib/utils/setting-defaults';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

interface BankAccountItem { name: string; bank: string; number: string; }

export default function AccountingSettings({ settings, onSave, saving }: TabProps) {
  const [businessInfo, setBusinessInfo] = useState({
    company: '', registration_number: '', representative: '',
    address: '', phone: '', business_type: '', business_item: '',
  });
  const [vatRate, setVatRate] = useState(10);
  const [expenseCategories, setExpenseCategories] = useState<string[]>([]);
  const [newCat, setNewCat] = useState('');
  const [cashflowIncomeCats, setCashflowIncomeCats] = useState<string[]>([]);
  const [newCfIncome, setNewCfIncome] = useState('');
  const [cashflowExpenseCats, setCashflowExpenseCats] = useState<string[]>([]);
  const [newCfExpense, setNewCfExpense] = useState('');
  const [bankAccounts, setBankAccounts] = useState<BankAccountItem[]>([]);
  const [taxType, setTaxType] = useState('general');
  const [revenueBasis, setRevenueBasis] = useState('sale_date');
  const [budgets, setBudgets] = useState<Record<string, number>>({});
  const [fiscalMonth, setFiscalMonth] = useState(1);

  useEffect(() => {
    setBusinessInfo(parse(settings['business.info'], businessInfo));
    setVatRate(parse(settings['accounting.vat_rate'], 10));
    setExpenseCategories(parse(settings['accounting.expense_categories'], DEFAULT_EXPENSE_CATEGORIES));
    setCashflowIncomeCats(parse(settings['accounting.cashflow_income_categories'], DEFAULT_CASHFLOW_INCOME_CATEGORIES));
    setCashflowExpenseCats(parse(settings['accounting.cashflow_expense_categories'], DEFAULT_CASHFLOW_EXPENSE_CATEGORIES));
    setBankAccounts(parse(settings['accounting.bank_accounts'], []));
    setTaxType(parse(settings['accounting.tax_type'], 'general'));
    setRevenueBasis(parse(settings['accounting.revenue_basis'], 'sale_date'));
    setBudgets(parse(settings['accounting.budgets'], {}));
    setFiscalMonth(parse(settings['accounting.fiscal_start_month'], 1));
  }, [settings]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSave = () => {
    onSave([
      { key: 'business.info', value: businessInfo },
      { key: 'accounting.vat_rate', value: vatRate },
      { key: 'accounting.expense_categories', value: expenseCategories },
      { key: 'accounting.cashflow_income_categories', value: cashflowIncomeCats },
      { key: 'accounting.cashflow_expense_categories', value: cashflowExpenseCats },
      { key: 'accounting.bank_accounts', value: bankAccounts },
      { key: 'accounting.tax_type', value: taxType },
      { key: 'accounting.revenue_basis', value: revenueBasis },
      { key: 'accounting.budgets', value: budgets },
      { key: 'accounting.fiscal_start_month', value: fiscalMonth },
    ]);
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">회계 설정</h2>

      {/* 1. 사업자 정보 (판매와 공유) */}
      <Field label="사업자 정보" desc="거래명세서·세금계산서·손익계산서에 공통 사용. (판매 설정과 공유)">
        <div className="grid grid-cols-2 gap-2">
          <input placeholder="상호" value={businessInfo.company} onChange={(e) => setBusinessInfo({ ...businessInfo, company: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="사업자번호" value={businessInfo.registration_number} onChange={(e) => setBusinessInfo({ ...businessInfo, registration_number: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="대표자" value={businessInfo.representative} onChange={(e) => setBusinessInfo({ ...businessInfo, representative: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="연락처" value={businessInfo.phone} onChange={(e) => setBusinessInfo({ ...businessInfo, phone: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="주소" value={businessInfo.address} onChange={(e) => setBusinessInfo({ ...businessInfo, address: e.target.value })}
            className="col-span-2 h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="업태" value={businessInfo.business_type} onChange={(e) => setBusinessInfo({ ...businessInfo, business_type: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
          <input placeholder="종목" value={businessInfo.business_item} onChange={(e) => setBusinessInfo({ ...businessInfo, business_item: e.target.value })}
            className="h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
        </div>
      </Field>

      {/* 2. 경비 카테고리 */}
      <Field label="경비 카테고리 목록" desc="경비 등록 시 선택 가능한 카테고리.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {expenseCategories.map((c, i) => (
            <span key={i} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
              {c}
              <button onClick={() => setExpenseCategories(expenseCategories.filter((_, j) => j !== i))} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
            </span>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newCat} onChange={(e) => setNewCat(e.target.value)} placeholder="새 카테고리"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            onKeyDown={(e) => { if (e.key === 'Enter' && newCat.trim()) { setExpenseCategories([...expenseCategories, newCat.trim()]); setNewCat(''); } }} />
          <button onClick={() => { if (newCat.trim()) { setExpenseCategories([...expenseCategories, newCat.trim()]); setNewCat(''); } }}
            className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 2-B. 입출금 — 입금 카테고리 */}
      <Field label="입출금 — 입금 카테고리" desc="입출금 페이지에서 입금 등록 시 선택 가능한 카테고리.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {cashflowIncomeCats.map((c, i) => (
            <span key={i} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
              {c}
              <button onClick={() => setCashflowIncomeCats(cashflowIncomeCats.filter((_, j) => j !== i))} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
            </span>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newCfIncome} onChange={(e) => setNewCfIncome(e.target.value)} placeholder="새 입금 카테고리"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            onKeyDown={(e) => { if (e.key === 'Enter' && newCfIncome.trim()) { setCashflowIncomeCats([...cashflowIncomeCats, newCfIncome.trim()]); setNewCfIncome(''); } }} />
          <button onClick={() => { if (newCfIncome.trim()) { setCashflowIncomeCats([...cashflowIncomeCats, newCfIncome.trim()]); setNewCfIncome(''); } }}
            className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 2-C. 입출금 — 출금 카테고리 */}
      <Field label="입출금 — 출금 카테고리" desc="입출금 페이지에서 출금 등록 시 선택 가능한 카테고리.">
        <div className="flex flex-wrap gap-1.5 mb-2">
          {cashflowExpenseCats.map((c, i) => (
            <span key={i} className="flex items-center gap-1 px-2 py-1 text-xs bg-neutral-100 rounded-lg">
              {c}
              <button onClick={() => setCashflowExpenseCats(cashflowExpenseCats.filter((_, j) => j !== i))} className="text-neutral-400 hover:text-red-500"><X size={12} /></button>
            </span>
          ))}
        </div>
        <div className="flex gap-2">
          <input value={newCfExpense} onChange={(e) => setNewCfExpense(e.target.value)} placeholder="새 출금 카테고리"
            className="flex-1 h-8 px-3 rounded-lg border border-neutral-200 text-sm"
            onKeyDown={(e) => { if (e.key === 'Enter' && newCfExpense.trim()) { setCashflowExpenseCats([...cashflowExpenseCats, newCfExpense.trim()]); setNewCfExpense(''); } }} />
          <button onClick={() => { if (newCfExpense.trim()) { setCashflowExpenseCats([...cashflowExpenseCats, newCfExpense.trim()]); setNewCfExpense(''); } }}
            className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 3. 부가세율 */}
      <Field label="부가세율" desc="">
        <div className="flex items-center gap-2">
          <input type="number" value={vatRate} onChange={(e) => setVatRate(Number(e.target.value))}
            className="w-20 h-9 px-3 rounded-lg border border-neutral-200 text-sm" min={0} max={100} />
          <span className="text-sm">%</span>
        </div>
      </Field>

      {/* 4. 입출금 계좌 */}
      <Field label="입출금 계좌 목록" desc="입출금 등록 시 계좌를 선택할 수 있습니다.">
        <div className="space-y-1.5 mb-2">
          {bankAccounts.map((acc, i) => (
            <div key={i} className="flex items-center gap-2 text-sm">
              <span>{acc.name}</span>
              <span className="text-neutral-500">{acc.bank} {acc.number}</span>
              <button onClick={() => setBankAccounts(bankAccounts.filter((_, j) => j !== i))}
                className="ml-auto text-neutral-400 hover:text-red-500"><X size={14} /></button>
            </div>
          ))}
        </div>
        <div className="flex gap-2">
          <input id="acc-name" placeholder="별칭 (사업자)" className="w-24 h-8 px-2 rounded-lg border border-neutral-200 text-sm" />
          <input id="acc-bank" placeholder="은행" className="w-20 h-8 px-2 rounded-lg border border-neutral-200 text-sm" />
          <input id="acc-num" placeholder="계좌번호" className="flex-1 h-8 px-2 rounded-lg border border-neutral-200 text-sm" />
          <button onClick={() => {
            const name = (document.getElementById('acc-name') as HTMLInputElement).value.trim();
            const bank = (document.getElementById('acc-bank') as HTMLInputElement).value.trim();
            const num = (document.getElementById('acc-num') as HTMLInputElement).value.trim();
            if (name && bank && num) {
              setBankAccounts([...bankAccounts, { name, bank, number: num }]);
              (document.getElementById('acc-name') as HTMLInputElement).value = '';
              (document.getElementById('acc-bank') as HTMLInputElement).value = '';
              (document.getElementById('acc-num') as HTMLInputElement).value = '';
            }
          }} className="px-2 rounded-lg bg-neutral-100 hover:bg-neutral-200"><Plus size={14} /></button>
        </div>
      </Field>

      {/* 5. 과세 구분 */}
      <Field label="과세 구분" desc="">
        <select value={taxType} onChange={(e) => setTaxType(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value="general">일반과세자</option>
          <option value="simplified">간이과세자</option>
        </select>
      </Field>

      {/* 6. 매출 인식 기준 */}
      <Field label="매출 인식 기준" desc="리포트에서 매출을 어느 시점 기준으로 잡을지.">
        <select value={revenueBasis} onChange={(e) => setRevenueBasis(e.target.value)}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          <option value="sale_date">판매일 기준</option>
          <option value="payment_date">결제일 기준</option>
          <option value="ship_date">출고일 기준</option>
        </select>
      </Field>

      {/* 7. 예산 관리 */}
      <Field label="월 예산 관리" desc="카테고리별 월 예산을 설정하면 경비 대비 잔여를 표시합니다.">
        <div className="space-y-1.5">
          {expenseCategories.map((cat) => (
            <div key={cat} className="flex items-center gap-2">
              <span className="text-sm w-20 truncate">{cat}</span>
              <input type="number" value={budgets[cat] || ''} onChange={(e) => setBudgets({ ...budgets, [cat]: Number(e.target.value) })}
                className="w-28 h-8 px-2 rounded-lg border border-neutral-200 text-sm" placeholder="0" step={10000} />
              <span className="text-xs text-neutral-500">원/월</span>
            </div>
          ))}
        </div>
      </Field>

      {/* 8. 회계 기간 */}
      <Field label="회계 기간 시작월" desc="">
        <select value={fiscalMonth} onChange={(e) => setFiscalMonth(Number(e.target.value))}
          className="h-9 px-3 rounded-lg border border-neutral-200 text-sm">
          {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => <option key={m} value={m}>{m}월</option>)}
        </select>
      </Field>

      {/* 9. 고정 경비 */}
      <Field label="고정 경비 자동 등록" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3">
          매월 반복 경비(임대료, 인건비 등)는 <span className="font-semibold">경비 페이지</span>에서 관리됩니다.
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
