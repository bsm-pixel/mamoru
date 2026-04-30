'use client';

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { Plus, Trash2, Wallet } from 'lucide-react';
import toast from 'react-hot-toast';

import { DEFAULT_EXPENSE_CATEGORIES } from '@/lib/utils/setting-defaults';
import { useSetting } from '@/hooks/use-settings';

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
  const [from, setFrom] = useState(new Date(now.getFullYear(), now.getMonth(), 1).toISOString().slice(0, 10));
  const [to, setTo] = useState(now.toISOString().slice(0, 10));

  // 폼 상태
  const [expenseDate, setExpenseDate] = useState(now.toISOString().slice(0, 10));
  const [category, setCategory] = useState<string>(CATEGORIES[0] || '기타');
  const [amount, setAmount] = useState('');
  const [memo, setMemo] = useState('');

  const { data, isLoading } = useQuery({
    queryKey: ['expenses', from, to],
    queryFn: async () => {
      const res = await fetch(`/api/expenses?from=${from}&to=${to}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        expenses: Array<{ id: string; expense_date: string; category: string; amount: number; memo: string | null; created_at: string }>;
        total: number;
        byCategory: Record<string, number>;
      }>;
    },
  });

  const createExpense = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/expenses', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ expense_date: expenseDate, category, amount: parseInt(amount), memo: memo.trim() || undefined }),
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

  const createRecurring = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/expenses/recurring', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'create', category: recCategory, amount: parseInt(recAmount), memo: recMemo.trim() || undefined }),
      });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => { toast.success('고정 경비 등록'); setRecAmount(''); setRecMemo(''); queryClient.invalidateQueries({ queryKey: ['recurring-expenses'] }); },
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
            {Object.entries(data.byCategory).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([cat, amt]) => (
              <Card key={cat}>
                <p className="text-xs text-neutral-500 mb-1">{cat}</p>
                <p className="text-base font-bold text-neutral-800">{formatKRW(amt)}</p>
              </Card>
            ))}
          </div>
        )}

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
                <input type="date" value={expenseDate} max={now.toISOString().slice(0, 10)}
                  onChange={(e) => setExpenseDate(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">카테고리</label>
                <div className="flex flex-wrap gap-1.5">
                  {CATEGORIES.map((c) => (
                    <button key={c} onClick={() => setCategory(c)}
                      className={`px-2.5 py-1 text-xs rounded-md border transition ${category === c ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}
                    >{c}</button>
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
                        <p className="text-xs text-neutral-400 mt-0.5">{formatDate(e.expense_date)}</p>
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
            <h3 className="text-sm font-bold text-indigo-black">고정 경비 (매월 반복)</h3>
            <Button
              size="sm"
              variant="secondary"
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
                  <span className="text-sm text-neutral-700 flex-1">{item.memo || item.category}</span>
                  <span className="text-sm font-bold">{formatKRW(item.amount)}</span>
                  <button
                    onClick={() => { if (confirm('고정 경비를 삭제하시겠습니까?')) deleteRecurring.mutate(item.id); }}
                    className="w-6 h-6 rounded flex items-center justify-center text-neutral-400 hover:text-red-500 hover:bg-red-50"
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
          <div className="flex gap-2 items-end">
            <select value={recCategory} onChange={(e) => setRecCategory(e.target.value)}
              className="h-9 px-2 rounded-lg border border-neutral-200 text-xs">
              {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
            <input type="number" value={recAmount} placeholder="금액"
              onChange={(e) => setRecAmount(e.target.value)}
              className="w-24 h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            <input type="text" value={recMemo} placeholder="메모"
              onChange={(e) => setRecMemo(e.target.value)}
              className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
            <Button size="sm" variant="secondary"
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
