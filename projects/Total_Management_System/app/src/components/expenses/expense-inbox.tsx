'use client';

/**
 * 메일에서 찾은 영수증 — 확인 대기 (2026-10-29, 마이그 166)
 * 매일 08시 크론이 bsm@mamoru.kr 영수증·세금계산서 메일을 읽어 올린다. 사장님은 금액·분류만 보고 [확인] 1번.
 * [확인] 때 고른 분류·고정/사용량은 거래처별로 기억 → 다음 메일부터 그 값으로 올라온다.
 */
import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { formatKRW } from '@/lib/utils/format';
import { Mail, RefreshCw, ExternalLink } from 'lucide-react';
import toast from 'react-hot-toast';
import type { ExpenseInboxItem } from '@/app/api/expenses/inbox/route';
import { PLAN_OPTIONS, toPlan, fromPlan, type ExpensePlan } from '@/lib/expenses/plan';

const DOC_LABEL: Record<ExpenseInboxItem['doc_type'], string> = {
  card: '카드결제',
  receipt: '영수증',
  invoice: '청구서',
  tax_invoice: '세금계산서',
};

interface ScanResult {
  ok: boolean; added: number; scanned: number; ignored: number; skipped: number;
  needsReauth?: boolean; notConnected?: boolean; apiDisabled?: boolean; error?: string;
}

function scanMessage(r: ScanResult): string {
  if (r.needsReauth) return '메일 읽기 권한이 없습니다 — 설정 → 구글 연결에서 [재연결] 한 번 눌러 주세요';
  if (r.notConnected) return '구글 계정이 연결돼 있지 않습니다 — 설정 → 구글 연결';
  if (r.apiDisabled) return '구글 클라우드에서 Gmail API가 꺼져 있습니다';
  return r.error ? `메일 확인 실패: ${r.error}` : '메일 확인 실패';
}

export function ExpenseInbox({ categories }: { categories: string[] }) {
  const queryClient = useQueryClient();
  const { data } = useQuery({
    queryKey: ['expense-inbox'],
    queryFn: async () => {
      const res = await fetch('/api/expenses/inbox');
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{ items: ExpenseInboxItem[] }>;
    },
  });

  const scan = useMutation({
    mutationFn: async (days: number) => {
      const res = await fetch('/api/expenses/inbox', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'scan', days }),
      });
      return res.json() as Promise<ScanResult>;
    },
    onSuccess: (r) => {
      if (!r.ok) { toast.error(scanMessage(r), { duration: 6000 }); return; }
      toast.success(r.added ? `영수증 ${r.added}건을 찾았습니다` : '새 영수증 메일이 없습니다');
      queryClient.invalidateQueries({ queryKey: ['expense-inbox'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  const items = data?.items || [];

  return (
    <Card>
      <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
        <h3 className="text-sm font-bold text-indigo-black flex items-center gap-2">
          <Mail size={16} />
          메일에서 찾은 영수증
          {items.length > 0 && <span className="px-1.5 py-0.5 rounded bg-neutral-900 text-white text-[11px]">확인 대기 {items.length}</span>}
        </h3>
        <div className="flex items-center gap-1">
          <button
            onClick={() => scan.mutate(120)}
            disabled={scan.isPending}
            className="h-8 px-2 text-xs text-neutral-500 hover:text-neutral-800 disabled:opacity-50"
            title="처음 한 번 — 지난 4개월 영수증 메일을 모읍니다"
          >지난 4개월 가져오기</button>
          <Button size="sm" variant="secondary" onClick={() => scan.mutate(3)} disabled={scan.isPending}>
            <RefreshCw size={13} className={scan.isPending ? 'animate-spin' : ''} />
            <span className="ml-1">메일 다시 확인</span>
          </Button>
        </div>
      </div>

      {items.length === 0 ? (
        <p className="text-xs text-neutral-400">
          확인할 영수증이 없습니다. 매일 아침 8시에 메일함(bsm@mamoru.kr)을 자동으로 확인합니다.
        </p>
      ) : (
        <div className="divide-y divide-neutral-100 -mx-1">
          {items.map((it) => <InboxRow key={it.id} item={it} categories={categories} />)}
        </div>
      )}
    </Card>
  );
}

function InboxRow({ item, categories }: { item: ExpenseInboxItem; categories: string[] }) {
  const queryClient = useQueryClient();
  const [amount, setAmount] = useState(item.amount != null ? String(item.amount) : '');
  const [date, setDate] = useState(item.expense_date || '');
  const [category, setCategory] = useState(item.category || '기타');
  // 167 고정 · 연간 · 사용량 (kind + billing_cycle)
  const [plan, setPlan] = useState<ExpensePlan>(item.kind ? toPlan(item.kind, item.billing_cycle) : 'variable');
  const cats = categories.includes(category) ? categories : [...categories, category];

  const act = useMutation({
    mutationFn: async (payload: Record<string, unknown>) => {
      const res = await fetch('/api/expenses/inbox', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ id: item.id, ...payload }),
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || '실패');
      return j;
    },
    onSuccess: (_r, payload) => {
      toast.success(payload.action === 'confirm' ? '경비로 등록했습니다' : '무시했습니다');
      queryClient.invalidateQueries({ queryKey: ['expense-inbox'] });
      queryClient.invalidateQueries({ queryKey: ['expenses'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  const amountNum = parseInt(amount, 10);
  const mailUrl = `https://mail.google.com/mail/u/0/#all/${item.gmail_id}`;

  return (
    <div className="px-1 py-3 flex flex-col lg:flex-row lg:items-center gap-2 lg:gap-3">
      {/* 1) 거래처 · 문서 · 메일 */}
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="text-sm font-semibold text-neutral-800 break-all">{item.vendor || '거래처 미상'}</span>
          <span className="px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-600 text-[10px] shrink-0">{DOC_LABEL[item.doc_type]}</span>
          <a href={mailUrl} target="_blank" rel="noreferrer" className="text-neutral-400 hover:text-neutral-700 shrink-0" title="메일 열기">
            <ExternalLink size={12} />
          </a>
        </div>
        <p className="text-[11px] text-neutral-400 truncate">{item.memo || item.subject}</p>
        {item.currency !== 'KRW' && item.amount_original != null && (
          // 2026-10-29 결제일 환율로 원화 추정(lib/expenses/fx.ts) — 환율 = 원화 ÷ 원금 (별도 칸 없이 역산)
          item.amount != null ? (
            <p className="text-[11px] text-neutral-600">
              {item.currency} {item.amount_original} × {Math.round(item.amount / Number(item.amount_original)).toLocaleString()}원 · 환율 추정 (명세서와 다르면 수정)
            </p>
          ) : (
            <p className="text-[11px] text-neutral-600">{item.currency} {item.amount_original} — 카드 청구 원화 금액을 넣어 주세요</p>
          )
        )}
        {item.amount == null && item.currency === 'KRW' && (
          <p className="text-[11px] text-neutral-600">금액이 첨부 PDF에만 있습니다 — 메일을 열어 확인 후 입력</p>
        )}
        {item.dup_hint && <p className="text-[11px] text-amber-700">이미 비슷한 경비가 있어요: {item.dup_hint}</p>}
        {plan === 'annual' && amountNum > 0 && (
          <p className="text-[11px] text-neutral-600">연간 결제 · 월 환산 {formatKRW(Math.round(amountNum / 12))} (장부엔 결제일 전액)</p>
        )}
      </div>

      {/* 2) 날짜 · 금액 · 분류 · 고정/사용량 */}
      <div className="grid grid-cols-2 sm:flex sm:items-center gap-1.5 shrink-0">
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
          className="h-8 px-2 rounded-md border border-neutral-200 text-xs min-w-0" />
        <input type="number" value={amount} placeholder="금액(원)" onChange={(e) => setAmount(e.target.value)}
          className="h-8 px-2 rounded-md border border-neutral-200 text-xs text-right sm:w-24 min-w-0" />
        <select value={category} onChange={(e) => setCategory(e.target.value)}
          className="h-8 px-1.5 rounded-md border border-neutral-200 text-xs min-w-0">
          {cats.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        <div className="flex rounded-md border border-neutral-200 overflow-hidden h-8 text-xs shrink-0">
          {PLAN_OPTIONS.map((o) => (
            <button key={o.value} onClick={() => setPlan(o.value)}
              className={`flex-1 px-2.5 whitespace-nowrap ${plan === o.value ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500'}`}>
              {o.short}
            </button>
          ))}
        </div>
      </div>

      {/* 3) 확인 / 무시 */}
      <div className="flex items-center gap-1 shrink-0">
        <button
          onClick={() => act.mutate({ action: 'confirm', amount: amountNum, expense_date: date, category, ...fromPlan(plan) })}
          disabled={!amountNum || act.isPending}
          className="h-8 px-3 rounded-md bg-neutral-900 text-white text-xs font-medium hover:bg-neutral-700 disabled:opacity-40 flex-1 lg:flex-none"
        >확인 {amountNum ? formatKRW(amountNum) : ''}</button>
        <button onClick={() => act.mutate({ action: 'ignore' })} disabled={act.isPending}
          className="h-8 px-2 rounded-md text-xs text-neutral-500 hover:bg-neutral-100">무시</button>
        <button
          onClick={() => { if (confirm(`「${item.vendor}」 메일은 앞으로 목록에 올리지 않을까요?`)) act.mutate({ action: 'ignore', always: true }); }}
          disabled={act.isPending}
          className="h-8 px-2 rounded-md text-[11px] text-neutral-400 hover:bg-neutral-100"
          title="이 거래처 메일은 앞으로 자동 무시"
        >늘 무시</button>
      </div>
    </div>
  );
}
