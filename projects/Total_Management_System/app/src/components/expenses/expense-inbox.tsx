'use client';

/**
 * 메일에서 찾은 영수증 — 확인 대기 (2026-10-29, 마이그 166)
 * 매일 08시 크론이 bsm@mamoru.kr 영수증·세금계산서 메일을 읽어 올린다. 사장님은 금액·분류만 보고 [확인] 1번.
 * [확인] 때 고른 분류·고정/사용량은 거래처별로 기억 → 다음 메일부터 그 값으로 올라온다.
 * 2026-10-29 카드 이용내역 엑셀도 같은 대기함으로 (api/expenses/card-import) + 기억된 거래처 한 번에 확인.
 */
import { useRef, useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { formatKRW } from '@/lib/utils/format';
import { Mail, RefreshCw, ExternalLink, Upload, CheckCheck } from 'lucide-react';
import toast from 'react-hot-toast';
import type { ExpenseInboxItem } from '@/app/api/expenses/inbox/route';
import { PLAN_OPTIONS, toPlan, fromPlan, type ExpensePlan } from '@/lib/expenses/plan';
import { PROOF_OPTIONS, proofFromDocType, toProof, type ProofType } from '@/lib/expenses/proof';

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
      return res.json() as Promise<{ items: ExpenseInboxItem[]; knownCount?: number }>;
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

  // 카드 이용내역 엑셀 올리기
  const fileRef = useRef<HTMLInputElement>(null);
  const upload = useMutation({
    mutationFn: async (file: File) => {
      const fd = new FormData();
      fd.append('file', file);
      const res = await fetch('/api/expenses/card-import', { method: 'POST', body: fd });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || '올리기 실패');
      return j as { rows: number; added: number; ignored: number; existed: number; cancelled: number };
    },
    onSuccess: (r) => {
      const extra = [r.existed ? `이미 올린 ${r.existed}건` : '', r.ignored ? `늘 무시 ${r.ignored}건` : '', r.cancelled ? `취소 ${r.cancelled}건` : ''].filter(Boolean).join(' · ');
      toast.success(`카드 내역 ${r.added}건을 올렸습니다${extra ? ` (${extra} 제외)` : ''}`, { duration: 6000 });
      queryClient.invalidateQueries({ queryKey: ['expense-inbox'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err), { duration: 6000 }),
  });

  // 기억된 거래처 한 번에 확인
  const confirmKnown = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/expenses/inbox', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'confirm_known' }),
      });
      const j = await res.json();
      if (!res.ok) throw new Error(j.error || '실패');
      return j as { confirmed: number };
    },
    onSuccess: (r) => {
      toast.success(`${r.confirmed}건을 경비로 등록했습니다`);
      queryClient.invalidateQueries({ queryKey: ['expense-inbox'] });
      queryClient.invalidateQueries({ queryKey: ['expenses'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  const [shown, setShown] = useState(15); // 카드 1년치 수백 건 — 15건씩
  const items = data?.items || [];
  const knownCount = data?.knownCount || 0;

  return (
    <Card>
      <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
        <h3 className="text-sm font-bold text-indigo-black flex items-center gap-2">
          <Mail size={16} />
          메일·카드에서 찾은 경비
          {items.length > 0 && <span className="px-1.5 py-0.5 rounded bg-neutral-900 text-white text-[11px]">확인 대기 {items.length}</span>}
        </h3>
        <div className="flex flex-wrap items-center gap-1">
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="hidden"
            onChange={(e) => { const f = e.target.files?.[0]; if (f) upload.mutate(f); e.target.value = ''; }} />
          <Button size="sm" variant="secondary" onClick={() => fileRef.current?.click()} disabled={upload.isPending}
            title="홈택스 「사업용 신용카드 사용내역」 또는 카드사 이용내역 엑셀">
            <Upload size={13} />
            <span className="ml-1">{upload.isPending ? '올리는 중...' : '카드 내역 올리기'}</span>
          </Button>
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

      {knownCount > 0 && (
        // 기억된 거래처 = 한 번 [확인]해서 분류를 정한 곳. 중복 의심 건은 빼고 한 번에
        <div className="flex flex-wrap items-center justify-between gap-2 mb-2 px-3 py-2 rounded-md bg-neutral-50">
          <p className="text-xs text-neutral-600">전에 확인한 거래처 <b>{knownCount}건</b>은 같은 분류로 한 번에 등록할 수 있어요</p>
          <Button size="sm" onClick={() => { if (confirm(`${knownCount}건을 기억된 분류 그대로 경비로 등록할까요?`)) confirmKnown.mutate(); }} disabled={confirmKnown.isPending}>
            <CheckCheck size={13} />
            <span className="ml-1">{confirmKnown.isPending ? '등록 중...' : `${knownCount}건 한 번에 확인`}</span>
          </Button>
        </div>
      )}

      {items.length === 0 ? (
        <p className="text-xs text-neutral-400">
          확인할 경비가 없습니다. 매일 아침 8시에 메일함(bsm@mamoru.kr)을 자동으로 확인합니다.
          카드로 쓴 것은 홈택스 「사업용 신용카드 사용내역」 엑셀을 [카드 내역 올리기]로 올려 주세요.
        </p>
      ) : (
        <div className="divide-y divide-neutral-100 -mx-1">
          {items.slice(0, shown).map((it) => <InboxRow key={it.id} item={it} categories={categories} />)}
          {items.length > shown && (
            <button onClick={() => setShown((n) => n + 15)} className="w-full py-2.5 text-xs text-neutral-500 hover:bg-neutral-50">
              더 보기 ({items.length - shown}건 남음)
            </button>
          )}
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
  // 168 증빙 — 거래처 기억·메일 종류로 채워져 옴 (마이그 전 행은 메일 종류 기본값)
  const [proof, setProof] = useState<ProofType>(toProof(item.proof_type) ?? proofFromDocType(item.doc_type));
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
  const isCardRow = item.gmail_id.startsWith('card:'); // 카드 이용내역 엑셀 줄 — 열 메일이 없음

  return (
    <div className="px-1 py-3 flex flex-col lg:flex-row lg:items-center gap-2 lg:gap-3">
      {/* 1) 거래처 · 문서 · 메일 */}
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-1.5 flex-wrap">
          <span className="text-sm font-semibold text-neutral-800 break-all">{item.vendor || '거래처 미상'}</span>
          <span className="px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-600 text-[10px] shrink-0">{DOC_LABEL[item.doc_type]}</span>
          {isCardRow ? (
            <span className="text-[10px] text-neutral-400 shrink-0">카드 내역</span>
          ) : (
            <a href={mailUrl} target="_blank" rel="noreferrer" className="text-neutral-400 hover:text-neutral-700 shrink-0" title="메일 열기">
              <ExternalLink size={12} />
            </a>
          )}
        </div>
        <p className="text-[11px] text-neutral-400 truncate">{isCardRow ? (item.memo || '') : (item.memo || item.subject)}</p>
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

      {/* 2) 날짜 · 금액 · 분류 · 증빙 · 고정/사용량 */}
      <div className="grid grid-cols-2 sm:flex sm:items-center gap-1.5 shrink-0">
        <input type="date" value={date} onChange={(e) => setDate(e.target.value)}
          className="h-8 px-2 rounded-md border border-neutral-200 text-xs min-w-0" />
        <input type="number" value={amount} placeholder="금액(원)" onChange={(e) => setAmount(e.target.value)}
          className="h-8 px-2 rounded-md border border-neutral-200 text-xs text-right sm:w-24 min-w-0" />
        <select value={category} onChange={(e) => setCategory(e.target.value)}
          className="h-8 px-1.5 rounded-md border border-neutral-200 text-xs min-w-0">
          {cats.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        <select value={proof} onChange={(e) => setProof(e.target.value as ProofType)} aria-label="증빙"
          title={PROOF_OPTIONS.find((o) => o.value === proof)?.hint}
          className="h-8 px-1.5 rounded-md border border-neutral-200 text-xs min-w-0">
          {PROOF_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
        </select>
        <div className="col-span-2 sm:col-auto flex rounded-md border border-neutral-200 overflow-hidden h-8 text-xs shrink-0">
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
          onClick={() => act.mutate({ action: 'confirm', amount: amountNum, expense_date: date, category, ...fromPlan(plan), proof_type: proof })}
          disabled={!amountNum || act.isPending}
          className="h-8 px-3 rounded-md bg-neutral-900 text-white text-xs font-medium hover:bg-neutral-700 disabled:opacity-40 flex-1 lg:flex-none"
        >확인 {amountNum ? formatKRW(amountNum) : ''}</button>
        <button onClick={() => act.mutate({ action: 'ignore' })} disabled={act.isPending}
          className="h-8 px-2 rounded-md text-xs text-neutral-500 hover:bg-neutral-100">무시</button>
        <button
          onClick={() => { if (confirm(`「${item.vendor}」은(는) 앞으로 목록에 올리지 않을까요? (지금 대기 중인 같은 거래처 건도 함께 무시)`)) act.mutate({ action: 'ignore', always: true }); }}
          disabled={act.isPending}
          className="h-8 px-2 rounded-md text-[11px] text-neutral-400 hover:bg-neutral-100"
          title="이 거래처 메일은 앞으로 자동 무시"
        >늘 무시</button>
      </div>
    </div>
  );
}
