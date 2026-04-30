'use client';

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { formatKRW, formatDate } from '@/lib/utils/format';
import { ArrowDownCircle, ArrowUpCircle, Trash2, TrendingUp, TrendingDown } from 'lucide-react';
import toast from 'react-hot-toast';
import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CASHFLOW_INCOME_CATEGORIES, DEFAULT_CASHFLOW_EXPENSE_CATEGORIES } from '@/lib/utils/setting-defaults';

export default function CashflowPage() {
  const queryClient = useQueryClient();
  const now = new Date();
  const [from, setFrom] = useState(new Date(now.getFullYear(), now.getMonth(), 1).toISOString().slice(0, 10));
  const [to, setTo] = useState(now.toISOString().slice(0, 10));

  // 설정값 (미저장 시 기본값 fallback — 기존 hard-coded와 동일)
  const INCOME_CATEGORIES = useSetting<string[]>('accounting.cashflow_income_categories', DEFAULT_CASHFLOW_INCOME_CATEGORIES);
  const EXPENSE_CATEGORIES = useSetting<string[]>('accounting.cashflow_expense_categories', DEFAULT_CASHFLOW_EXPENSE_CATEGORIES);

  const [txDate, setTxDate] = useState(now.toISOString().slice(0, 10));
  const [txType, setTxType] = useState<'income' | 'expense'>('income');
  const [txCategory, setTxCategory] = useState(INCOME_CATEGORIES[0] || '매출입금');
  const [txAmount, setTxAmount] = useState('');
  const [txMemo, setTxMemo] = useState('');

  const { data, isLoading } = useQuery({
    queryKey: ['cashflow', from, to],
    queryFn: async () => {
      const res = await fetch(`/api/cashflow?from=${from}&to=${to}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        transactions: Array<{ id: string; transaction_date: string; type: string; category: string; amount: number; memo: string | null }>;
        summary: { income: number; expense: number; net: number };
      }>;
    },
  });

  const createTx = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/cashflow', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ transaction_date: txDate, type: txType, category: txCategory, amount: parseInt(txAmount), memo: txMemo.trim() || undefined }),
      });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => { toast.success('등록 완료'); setTxAmount(''); setTxMemo(''); queryClient.invalidateQueries({ queryKey: ['cashflow'] }); },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  const deleteTx = useMutation({
    mutationFn: async (id: string) => { await fetch(`/api/cashflow?id=${id}`, { method: 'DELETE' }); },
    onSuccess: () => { toast.success('삭제됨'); queryClient.invalidateQueries({ queryKey: ['cashflow'] }); },
  });

  const categories = txType === 'income' ? INCOME_CATEGORIES : EXPENSE_CATEGORIES;

  return (
    <>
      <Topbar title="입출금 관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 요약 카드 */}
        {data && (
          <div className="grid grid-cols-3 gap-3">
            <Card>
              <div className="flex items-center gap-1.5 text-xs text-green-600 mb-1"><ArrowDownCircle size={12} /> 입금</div>
              <p className="text-lg font-bold text-green-600">{formatKRW(data.summary.income)}</p>
            </Card>
            <Card>
              <div className="flex items-center gap-1.5 text-xs text-red-500 mb-1"><ArrowUpCircle size={12} /> 출금</div>
              <p className="text-lg font-bold text-red-500">{formatKRW(data.summary.expense)}</p>
            </Card>
            <Card>
              <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
                {data.summary.net >= 0 ? <TrendingUp size={12} /> : <TrendingDown size={12} />} 순이익
              </div>
              <p className={`text-lg font-bold ${data.summary.net >= 0 ? 'text-neutral-800' : 'text-red-500'}`}>{formatKRW(data.summary.net)}</p>
            </Card>
          </div>
        )}

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
          {/* 등록 폼 */}
          <Card>
            <h3 className="text-sm font-bold text-indigo-black mb-3">입출금 등록</h3>
            <div className="space-y-3">
              {/* 입금/출금 토글 */}
              <div className="flex gap-2">
                <button onClick={() => { setTxType('income'); setTxCategory(INCOME_CATEGORIES[0] || ''); }}
                  className={`flex-1 py-2 rounded-lg text-sm font-medium transition ${txType === 'income' ? 'bg-green-600 text-white' : 'bg-neutral-100 text-neutral-500'}`}>
                  입금
                </button>
                <button onClick={() => { setTxType('expense'); setTxCategory(EXPENSE_CATEGORIES[0] || ''); }}
                  className={`flex-1 py-2 rounded-lg text-sm font-medium transition ${txType === 'expense' ? 'bg-red-500 text-white' : 'bg-neutral-100 text-neutral-500'}`}>
                  출금
                </button>
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">날짜</label>
                <input type="date" value={txDate} onChange={(e) => setTxDate(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">카테고리</label>
                <select value={txCategory} onChange={(e) => setTxCategory(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm">
                  {categories.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">금액</label>
                <input type="number" value={txAmount} placeholder="0" onChange={(e) => setTxAmount(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">메모</label>
                <input type="text" value={txMemo} placeholder="상세 내용" onChange={(e) => setTxMemo(e.target.value)}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm placeholder:text-neutral-400" />
              </div>
              <Button className="w-full" onClick={() => createTx.mutate()}
                disabled={!txAmount || parseInt(txAmount) <= 0 || createTx.isPending}>
                {createTx.isPending ? '등록 중...' : '등록'}
              </Button>
            </div>
          </Card>

          {/* 내역 목록 */}
          <div className="lg:col-span-2">
            <div className="flex items-center gap-2 mb-3">
              <input type="date" value={from} onChange={(e) => setFrom(e.target.value)} className="h-8 px-2 rounded-lg border border-neutral-200 text-xs" />
              <span className="text-xs text-neutral-400">~</span>
              <input type="date" value={to} onChange={(e) => setTo(e.target.value)} className="h-8 px-2 rounded-lg border border-neutral-200 text-xs" />
            </div>
            <Card padding={false}>
              {isLoading ? (
                <div className="p-6 text-center text-sm text-neutral-400">로딩중...</div>
              ) : !data?.transactions.length ? (
                <div className="p-6 text-center text-sm text-neutral-400">입출금 내역이 없습니다</div>
              ) : (
                <div className="divide-y divide-neutral-100">
                  {data.transactions.map((tx) => (
                    <div key={tx.id} className="flex items-center gap-3 px-4 py-3">
                      <div className={`w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${tx.type === 'income' ? 'bg-green-50 text-green-600' : 'bg-red-50 text-red-500'}`}>
                        {tx.type === 'income' ? <ArrowDownCircle size={14} /> : <ArrowUpCircle size={14} />}
                      </div>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2">
                          <span className="text-xs px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-600">{tx.category}</span>
                          {tx.memo && <span className="text-sm text-neutral-700 truncate">{tx.memo}</span>}
                        </div>
                        <p className="text-xs text-neutral-400 mt-0.5">{formatDate(tx.transaction_date)}</p>
                      </div>
                      <p className={`text-sm font-bold shrink-0 ${tx.type === 'income' ? 'text-green-600' : 'text-red-500'}`}>
                        {tx.type === 'income' ? '+' : '-'}{formatKRW(tx.amount)}
                      </p>
                      <button onClick={() => { if (confirm('삭제?')) deleteTx.mutate(tx.id); }}
                        className="shrink-0 w-7 h-7 rounded flex items-center justify-center text-neutral-400 hover:text-red-500 hover:bg-red-50">
                        <Trash2 size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </Card>
          </div>
        </div>
      </div>
    </>
  );
}
