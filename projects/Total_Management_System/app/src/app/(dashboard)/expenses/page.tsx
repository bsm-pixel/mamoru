'use client';

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { formatKRW, formatDate, toLocalDateString } from '@/lib/utils/format';
import { Plus, Trash2, Wallet } from 'lucide-react';
import toast from 'react-hot-toast';

import { DEFAULT_EXPENSE_CATEGORIES } from '@/lib/utils/setting-defaults';
import { useSetting } from '@/hooks/use-settings';
import { ExpenseInbox } from '@/components/expenses/expense-inbox';
import { PLAN_OPTIONS, fromPlan, planLabel, type ExpensePlan } from '@/lib/expenses/plan';
import { PROOF_OPTIONS, proofNeedsKeeping, type ProofType } from '@/lib/expenses/proof';

const CATEGORY_COLOR: Record<string, string> = {
  택배비: 'bg-blue-100 text-blue-700',
  포장재: 'bg-green-100 text-green-700',
  교통비: 'bg-yellow-100 text-yellow-700',
  사무용품: 'bg-purple-100 text-purple-700',
  식대: 'bg-orange-100 text-orange-700',
  소모품: 'bg-neutral-100 text-neutral-600',
  기타: 'bg-neutral-100 text-neutral-600',
};

export default function ExpensesPage() {
  const queryClient = useQueryClient();
  // 075: 설정에서 동적으로 카테고리 읽기 (사장님이 추가한 카테고리 자동 반영)
  const CATEGORIES = useSetting<string[]>('accounting.expense_categories', DEFAULT_EXPENSE_CATEGORIES);
  const now = new Date();
  const [from, setFrom] = useState(toLocalDateString(new Date(now.getFullYear(), now.getMonth(), 1)));
  const [to, setTo] = useState(toLocalDateString(now));

  // 폼 상태
  const [expenseDate, setExpenseDate] = useState(toLocalDateString(now));
  const [category, setCategory] = useState<string>(CATEGORIES[0] || '기타');
  const [amount, setAmount] = useState('');
  const [memo, setMemo] = useState('');
  const [plan, setPlan] = useState<ExpensePlan>('variable'); // 166 고정/사용량 + 167 연간
  const [proof, setProof] = useState<ProofType>('card'); // 168 증빙 — 사업자카드(7620)가 기본

  const { data, isLoading } = useQuery({
    queryKey: ['expenses', from, to],
    queryFn: async () => {
      const res = await fetch(`/api/expenses?from=${from}&to=${to}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        expenses: Array<{ id: string; expense_date: string; category: string; amount: number; memo: string | null; created_at: string; kind: 'fixed' | 'variable' | null; billing_cycle: 'annual' | 'monthly' | null; proof_type: ProofType | null; source_ref: string | null; receipt_url: string | null }>;
        total: number;
        byCategory: Record<string, number>;
        byKind: { fixed: number; variable: number; none: number; annual: number };
        fixedMonthly: number; // 167 월 환산 고정비 (연간 ÷ 12 반영, 참고값)
        fixedMonths: number;
      }>;
    },
  });

  const createExpense = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/expenses', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ expense_date: expenseDate, category, amount: parseInt(amount), memo: memo.trim() || undefined, ...fromPlan(plan), proof_type: proof }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: () => {
      toast.success('경비 등록 완료');
      setAmount('');
      setMemo('');
      queryClient.invalidateQueries({ queryKey: ['expenses'] });
    },
    onError: (err) => toast.error('등록 실패: ' + (err instanceof Error ? err.message : String(err))),
  });

  // 168 증빙 — 목록에서 바로 고침 (메일로 들어온 옛 건·미정 정리용)
  const updateProof = useMutation({
    mutationFn: async ({ id, proof_type }: { id: string; proof_type: ProofType | null }) => {
      const res = await fetch('/api/expenses', {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id, proof_type }),
      });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => queryClient.invalidateQueries({ queryKey: ['expenses'] }),
    onError: (err) => toast.error('증빙 저장 실패: ' + (err instanceof Error ? err.message : String(err))),
  });

  const deleteExpense = useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/expenses?id=${id}`, { method: 'DELETE' });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => {
      toast.success('삭제됨');
      queryClient.invalidateQueries({ queryKey: ['expenses'] });
    },
  });

  // 고정 경비
  const { data: recurringData } = useQuery({
    queryKey: ['recurring-expenses'],
    queryFn: async () => {
      const res = await fetch('/api/expenses/recurring');
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ items: Array<{ id: string; category: string; amount: number; memo: string | null }> }>;
    },
  });

  const [recCategory, setRecCategory] = useState<string>(CATEGORIES.find((c) => c === '임대료') || CATEGORIES[0] || '기타');
  const [recAmount, setRecAmount] = useState('');
  const [recMemo, setRecMemo] = useState('');
  // 169 시작 월 — 그 달부터 이번 달까지 빠진 달을 한 번에 채움 (기본: 올해 1월)
  const [recFrom, setRecFrom] = useState(`${now.getFullYear()}-01`);
  const [recProof, setRecProof] = useState<ProofType>('none');

  const createRecurring = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/expenses/recurring', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'create', category: recCategory, amount: parseInt(recAmount), memo: recMemo.trim() || undefined, from: recFrom, proof_type: recProof }),
      });
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ generated: number; skipped: number }>;
    },
    onSuccess: (r) => {
      toast.success(`고정 경비 등록 — ${r.generated}개월치 경비 생성${r.skipped ? ` (이미 있던 ${r.skipped}개월 건너뜀)` : ''}`, { duration: 6000 });
      setRecAmount(''); setRecMemo('');
      queryClient.invalidateQueries({ queryKey: ['recurring-expenses'] });
      queryClient.invalidateQueries({ queryKey: ['expenses'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  const deleteRecurring = useMutation({
    mutationFn: async (id: string) => {
      await fetch('/api/expenses/recurring', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'delete', id }) });
    },
    onSuccess: () => { toast.success('삭제됨'); queryClient.invalidateQueries({ queryKey: ['recurring-expenses'] }); },
  });

  const generateMonthly = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/expenses/recurring', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'generate' }) });
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    onSuccess: (data) => {
      if (data.generated > 0) { toast.success(`이번달 고정 경비 ${data.generated}건 등록`); queryClient.invalidateQueries({ queryKey: ['expenses'] }); }
      else toast(data.message || '이미 등록됨');
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  return (
    <>
      <Topbar title="경비 관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 이번달 합계 + 카테고리별 */}
        {data && (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <Card>
              <div className="flex items-center gap-2 mb-1">
                <Wallet size={14} className="text-neutral-500" />
                <p className="text-xs text-neutral-500">이번달 경비</p>
              </div>
              <p className="text-xl font-bold text-red-600">{formatKRW(data.total)}</p>
              <p className="text-[11px] text-neutral-400">{data.expenses.length}건</p>
            </Card>
            {/* 166: 고정(구독·임대) / 사용량(택배·부자재) */}
            <Card>
              <p className="text-xs text-neutral-500 mb-1">고정 지출</p>
              <p className="text-base font-bold text-neutral-800">{formatKRW(data.byKind?.fixed || 0)}</p>
              {/* 167: 연간 결제가 결제한 달에만 크게 잡히는 것 보정 — 월 환산은 참고값, 장부는 결제일 전액 */}
              <p className="text-[11px] text-neutral-400">
                {data.fixedMonthly !== undefined && data.fixedMonthly !== (data.byKind?.fixed || 0)
                  ? `월 환산 ${formatKRW(data.fixedMonthly)}${data.fixedMonths > 1 ? ` (${data.fixedMonths}개월)` : ''}`
                  : '구독·임대 등 매달 나가는 돈'}
              </p>
            </Card>
            <Card>
              <p className="text-xs text-neutral-500 mb-1">사용량 지출</p>
              <p className="text-base font-bold text-neutral-800">{formatKRW(data.byKind?.variable || 0)}</p>
              <p className="text-[11px] text-neutral-400">택배·부자재 등 쓴 만큼{data.byKind?.none ? ` · 미분류 ${formatKRW(data.byKind.none)}` : ''}</p>
            </Card>
            {Object.entries(data.byCategory).sort((a, b) => b[1] - a[1]).slice(0, 1).map(([cat, amt]) => (
              <Card key={cat}>
                <p className="text-xs text-neutral-500 mb-1">가장 큰 항목 · {cat}</p>
                <p className="text-base font-bold text-neutral-800">{formatKRW(amt)}</p>
              </Card>
            ))}
          </div>
        )}

        {/* 166: 메일에서 찾은 영수증 — 확인 대기 */}
        <ExpenseInbox categories={CATEGORIES} />

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          {/* 좌측: 등록 폼 */}
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3 flex items-center gap-2">
              <Plus size={16} />
              경비 등록
            </h3>
            <div className="space-y-3">
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">날짜</label>
                <input type="date" value={expenseDate} max={toLocalDateString(now)}
                  onChange={(e) => setExpenseDate(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">카테고리</label>
                <div className="flex flex-wrap gap-1.5">
                  {CATEGORIES.map((c) => (
                    <button key={c} onClick={() => setCategory(c)}
                      className={`h-8 lg:h-auto px-2.5 py-1 text-xs rounded-md border transition ${category === c /* 2026-10-19 모바일 탭 타겟 32px */ ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                    >{c}</button>
                  ))}
                </div>
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">구분</label>
                <div className="flex rounded-md border border-neutral-200 overflow-hidden h-8 text-xs">
                  {PLAN_OPTIONS.map((o) => (
                    <button key={o.value} onClick={() => setPlan(o.value)}
                      className={`flex-1 ${plan === o.value ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500'}`}>
                      {o.long}
                    </button>
                  ))}
                </div>
              </div>
              <div>
                {/* 168 증빙 — 종소세 간편장부 「결제·증빙」 칸. 영수증만·없음은 「증빙 챙길 것」에 올라감 */}
                <label className="text-xs text-neutral-500 mb-1 block">증빙</label>
                <div className="flex flex-wrap gap-1.5">
                  {PROOF_OPTIONS.map((o) => (
                    <button key={o.value} onClick={() => setProof(o.value)} title={o.hint}
                      className={`h-8 lg:h-auto px-2.5 py-1 text-xs rounded-md border transition ${proof === o.value ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                    >{o.label}</button>
                  ))}
                </div>
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">금액</label>
                <input type="number" value={amount} placeholder="0"
                  onChange={(e) => setAmount(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">메모 (선택)</label>
                <input type="text" value={memo} placeholder="예: 롯데택배 4월 수거비"
                  onChange={(e) => setMemo(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
              </div>
              <Button
                className="w-full"
                onClick={() => createExpense.mutate()}
                disabled={!amount || parseInt(amount) <= 0 || createExpense.isPending}
              >
                {createExpense.isPending ? '등록 중...' : '경비 등록'}
              </Button>
            </div>
          </Card>

          {/* 우측: 목록 */}
          <div className="lg:col-span-2">
            {/* 기간 필터 */}
            <div className="flex items-center gap-2 mb-3">
              <input type="date" value={from} onChange={(e) => setFrom(e.target.value)}
                className="h-8 px-2 rounded-lg border border-neutral-200 text-xs" />
              <span className="text-xs text-neutral-400">~</span>
              <input type="date" value={to} onChange={(e) => setTo(e.target.value)}
                className="h-8 px-2 rounded-lg border border-neutral-200 text-xs" />
            </div>

            <Card padding={false}>
              {isLoading ? (
                <div className="p-6 text-center text-sm text-neutral-400">로딩중...</div>
              ) : !data?.expenses.length ? (
                <div className="p-6 text-center text-sm text-neutral-400">등록된 경비가 없습니다</div>
              ) : (
                <div className="divide-y divide-neutral-100">
                  {data.expenses.map((e) => (
                    <div key={e.id} className="flex items-center gap-3 px-4 py-3">
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2">
                          <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${CATEGORY_COLOR[e.category] || CATEGORY_COLOR['기타']}`}>
                            {e.category}
                          </span>
                          {e.memo && <span className="text-sm text-neutral-700 truncate">{e.memo}</span>}
                        </div>
                        <p className="text-xs text-neutral-400 mt-0.5">
                          {formatDate(e.expense_date)}
                          {e.kind && <span className="ml-1.5">· {planLabel(e.kind, e.billing_cycle)}</span>}
                          {e.source_ref?.startsWith('mail:') && <span className="ml-1.5">· 메일</span>}
                          {e.source_ref?.startsWith('card:') && <span className="ml-1.5">· 카드 내역</span>}
                          {/* 보관한 영수증 원본 (lib/expenses/receipt-archive) — 여러 개면 첫 번째 */}
                          {e.receipt_url && <a href={`/api/expenses/receipt?id=${e.id}`} target="_blank" rel="noreferrer" className="ml-1.5 underline hover:text-neutral-700">· 영수증</a>}
                          {/* 168 증빙 — 바로 고침. 홈택스 자동이 아닌 것(영수증만·없음·미정)은 진하게 */}
                          <span className="ml-1.5">·</span>
                          <select value={e.proof_type ?? ''} aria-label="증빙"
                            onChange={(ev) => updateProof.mutate({ id: e.id, proof_type: (ev.target.value || null) as ProofType | null })}
                            className={`ml-1 bg-transparent text-xs cursor-pointer hover:underline ${proofNeedsKeeping(e.proof_type) ? 'text-amber-700 font-medium' : 'text-neutral-400'}`}>
                            {!e.proof_type && <option value="">증빙 미정</option>}
                            {PROOF_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                          </select>
                        </p>
                      </div>
                      <p className="text-sm font-bold text-neutral-800 shrink-0">{formatKRW(e.amount)}</p>
                      <button
                        onClick={() => { if (confirm('삭제하시겠습니까?')) deleteExpense.mutate(e.id); }}
                        className="shrink-0 w-7 h-7 rounded flex items-center justify-center text-neutral-400 hover:text-red-500 hover:bg-red-50 transition"
                      >
                        <Trash2 size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </div>
        </div>
        {/* 고정 경비 */}
        <Card>
          <div className="flex items-center justify-between mb-3">
            <div>
              <h3 className="text-sm font-bold text-indigo-black">고정 경비 (매월 반복)</h3>
              {/* 169 매일 크론이 이번 달 분을 자동 등록 — 버튼은 지금 바로 넣고 싶을 때만 */}
              <p className="text-[11px] text-neutral-400 mt-0.5">매달 1일자로 자동 등록됩니다 · 추가할 때 시작 월부터 빠진 달도 채웁니다</p>
            </div>
            <Button
              size="sm"
              variant="secondary"
              className="shrink-0 whitespace-nowrap"
              onClick={() => generateMonthly.mutate()}
              disabled={generateMonthly.isPending || !recurringData?.items.length}
            >
              {generateMonthly.isPending ? '생성 중...' : '이번달 일괄 등록'}
            </Button>
          </div>

          {/* 고정 경비 목록 */}
          {recurringData?.items && recurringData.items.length > 0 ? (
            <div className="space-y-2 mb-3">
              {recurringData.items.map((item) => (
                <div key={item.id} className="flex items-center gap-3 py-2 border-b border-neutral-50 last:border-0">
                  <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${CATEGORY_COLOR[item.category] || CATEGORY_COLOR['기타']}`}>
                    {item.category}
                  </span>
                  {/* 2026-10-19 모바일: 긴 메모 줄바꿈 허용·삭제버튼 탭 타겟 확대 */}
                  <span className="text-sm text-neutral-700 flex-1 min-w-0 break-words">{item.memo || item.category}</span>
                  <span className="text-sm font-bold shrink-0 whitespace-nowrap">{formatKRW(item.amount)}</span>
                  <button
                    onClick={() => { if (confirm('고정 경비를 삭제하시겠습니까?')) deleteRecurring.mutate(item.id); }}
                    className="shrink-0 w-9 h-9 lg:w-6 lg:h-6 rounded flex items-center justify-center text-neutral-400 hover:text-red-500 hover:bg-red-50"
                  >
                    <Trash2 size={12} />
                  </button>
                </div>
              ))}
              <p className="text-[11px] text-neutral-400">
                월 고정 합계: {formatKRW(recurringData.items.reduce((s, i) => s + i.amount, 0))}
              </p>
            </div>
          ) : (
            <p className="text-xs text-neutral-400 mb-3">등록된 고정 경비가 없습니다</p>
          )}

          {/* 고정 경비 등록 폼 */}
          {/* 2026-10-19 모바일: 2열 그리드(분류·금액 / 메모 / 추가)로 가로 넘침 방지, sm↑는 기존 한 줄 */}
          <div className="grid grid-cols-2 gap-2 sm:flex sm:items-end">
            <select value={recCategory} onChange={(e) => setRecCategory(e.target.value)}
              className="w-full sm:w-auto min-w-0 h-9 px-2 rounded-lg border border-neutral-200 text-xs">
              {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <input type="number" value={recAmount} placeholder="금액"
              onChange={(e) => setRecAmount(e.target.value)}
              className="w-full sm:w-24 min-w-0 h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            <input type="text" value={recMemo} placeholder="메모"
              onChange={(e) => setRecMemo(e.target.value)}
              className="col-span-2 w-full sm:w-auto sm:flex-1 min-w-0 h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
            <label className="flex items-center gap-1.5 min-w-0 text-[11px] text-neutral-500 whitespace-nowrap">
              시작
              <input type="month" value={recFrom} max={toLocalDateString(now).slice(0, 7)} onChange={(e) => setRecFrom(e.target.value)}
                aria-label="시작 월" className="w-full sm:w-36 min-w-0 h-9 px-2 rounded-lg border border-neutral-200 text-xs text-neutral-800" />
            </label>
            <select value={recProof} onChange={(e) => setRecProof(e.target.value as ProofType)} aria-label="고정 경비 증빙"
              className="w-full sm:w-auto min-w-0 h-9 px-2 rounded-lg border border-neutral-200 text-xs">
              {PROOF_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
            </select>
            <Button size="sm" variant="secondary" className="col-span-2 h-9 sm:h-8 shrink-0"
              onClick={() => createRecurring.mutate()}
              disabled={!recAmount || parseInt(recAmount) <= 0}>
              추가
            </Button>
          </div>
        </Card>
      </div>
    </>
  );
}
