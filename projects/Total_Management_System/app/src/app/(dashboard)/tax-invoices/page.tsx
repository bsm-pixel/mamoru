'use client';

import { useState } from 'react';
import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { formatKRW, formatDate, toLocalDateString } from '@/lib/utils/format';
import { FileText, Trash2, Plus } from 'lucide-react';
import toast from 'react-hot-toast';

export default function TaxInvoicesPage() {
  const queryClient = useQueryClient();
  const now = new Date();
  // 이번 분기 기본값
  const quarter = Math.floor(now.getMonth() / 3);
  const qStart = new Date(now.getFullYear(), quarter * 3, 1);
  const [from, setFrom] = useState(toLocalDateString(qStart));
  const [to, setTo] = useState(toLocalDateString(now));
  const [typeFilter, setTypeFilter] = useState<string>('all');
  const [showForm, setShowForm] = useState(false);

  // 폼
  const [invType, setInvType] = useState<'sales' | 'purchase'>('sales');
  const [issueDate, setIssueDate] = useState(toLocalDateString(now));
  const [counterparty, setCounterparty] = useState('');
  const [bizNo, setBizNo] = useState('');
  const [supplyAmount, setSupplyAmount] = useState('');
  const [taxAmount, setTaxAmount] = useState('');
  const [memo, setMemo] = useState('');

  const { data, isLoading } = useQuery({
    queryKey: ['tax-invoices', from, to, typeFilter],
    queryFn: async () => {
      const params = new URLSearchParams({ from, to });
      if (typeFilter !== 'all') params.set('type', typeFilter);
      const res = await fetch(`/api/tax-invoices?${params}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json() as Promise<{
        invoices: Array<{ id: string; invoice_type: string; issue_date: string; counterparty_name: string; counterparty_biz_no: string | null; supply_amount: number; tax_amount: number; total_amount: number; memo: string | null }>;
        summary: { salesTotal: number; purchaseTotal: number; salesTax: number; purchaseTax: number; netTax: number; count: number };
      }>;
    },
  });

  const createInvoice = useMutation({
    mutationFn: async () => {
      const res = await fetch('/api/tax-invoices', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          invoice_type: invType, issue_date: issueDate, counterparty_name: counterparty,
          counterparty_biz_no: bizNo.trim() || undefined,
          supply_amount: parseInt(supplyAmount) || 0,
          tax_amount: parseInt(taxAmount) || 0,
          memo: memo.trim() || undefined,
        }),
      });
      if (!res.ok) throw new Error(await res.text());
    },
    onSuccess: () => {
      toast.success('세금계산서 등록');
      setCounterparty(''); setBizNo(''); setSupplyAmount(''); setTaxAmount(''); setMemo(''); setShowForm(false);
      queryClient.invalidateQueries({ queryKey: ['tax-invoices'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });

  const deleteInvoice = useMutation({
    mutationFn: async (id: string) => { await fetch(`/api/tax-invoices?id=${id}`, { method: 'DELETE' }); },
    onSuccess: () => { toast.success('삭제됨'); queryClient.invalidateQueries({ queryKey: ['tax-invoices'] }); },
  });

  // 공급가 입력 시 세액 자동 계산
  const handleSupplyChange = (v: string) => {
    setSupplyAmount(v);
    const s = parseInt(v) || 0;
    setTaxAmount(String(Math.round(s * 0.1)));
  };

  return (
    <>
      <Topbar title="세금계산서 관리" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 요약 카드 */}
        {data && (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <Card>
              <p className="text-xs text-neutral-500 mb-1">매출 세금계산서</p>
              <p className="text-base font-bold">{formatKRW(data.summary.salesTotal)}</p>
              <p className="text-[11px] text-neutral-400">세액 {formatKRW(data.summary.salesTax)}</p>
            </Card>
            <Card>
              <p className="text-xs text-neutral-500 mb-1">매입 세금계산서</p>
              <p className="text-base font-bold">{formatKRW(data.summary.purchaseTotal)}</p>
              <p className="text-[11px] text-neutral-400">세액 {formatKRW(data.summary.purchaseTax)}</p>
            </Card>
            <Card>
              <p className="text-xs text-neutral-500 mb-1">납부 세액</p>
              <p className={`text-base font-bold ${data.summary.netTax >= 0 ? 'text-red-600' : 'text-green-600'}`}>
                {data.summary.netTax >= 0 ? formatKRW(data.summary.netTax) : `환급 ${formatKRW(Math.abs(data.summary.netTax))}`}
              </p>
            </Card>
            <Card>
              <p className="text-xs text-neutral-500 mb-1">총 건수</p>
              <p className="text-base font-bold">{data.summary.count}건</p>
            </Card>
          </div>
        )}

        {/* 필터 + 등록 */}
        <div className="flex items-center gap-2 flex-wrap">
          <input type="date" value={from} onChange={(e) => setFrom(e.target.value)} className="h-8 px-2 rounded-lg border border-neutral-200 text-xs" />
          <span className="text-xs text-neutral-400">~</span>
          <input type="date" value={to} onChange={(e) => setTo(e.target.value)} className="h-8 px-2 rounded-lg border border-neutral-200 text-xs" />
          <div className="flex gap-1 ml-2">
            {[{ key: 'all', label: '전체' }, { key: 'sales', label: '매출' }, { key: 'purchase', label: '매입' }].map((t) => (
              <button key={t.key} onClick={() => setTypeFilter(t.key)}
                className={`px-2.5 py-1 text-xs rounded-md border transition ${typeFilter === t.key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'}`}>
                {t.label}
              </button>
            ))}
          </div>
          <Button size="sm" className="ml-auto" onClick={() => setShowForm(!showForm)}>
            <Plus size={14} /> 등록
          </Button>
        </div>

        {/* 등록 폼 */}
        {showForm && (
          <Card>
            <h3 className="text-sm font-bold mb-3">세금계산서 등록</h3>
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">유형</label>
                <div className="flex gap-2">
                  <button onClick={() => setInvType('sales')} className={`flex-1 py-1.5 text-xs rounded-md border ${invType === 'sales' ? 'bg-blue-600 text-white' : 'bg-white text-neutral-500'}`}>매출</button>
                  <button onClick={() => setInvType('purchase')} className={`flex-1 py-1.5 text-xs rounded-md border ${invType === 'purchase' ? 'bg-orange-600 text-white' : 'bg-white text-neutral-500'}`}>매입</button>
                </div>
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">발행일</label>
                <input type="date" value={issueDate} onChange={(e) => setIssueDate(e.target.value)} className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">거래처명</label>
                <input type="text" value={counterparty} onChange={(e) => setCounterparty(e.target.value)} placeholder="거래처명" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">사업자번호</label>
                <input type="text" value={bizNo} onChange={(e) => setBizNo(e.target.value)} placeholder="000-00-00000" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">공급가액</label>
                <input type="number" value={supplyAmount} onChange={(e) => handleSupplyChange(e.target.value)} placeholder="0" className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">세액 (자동)</label>
                <input type="number" value={taxAmount} onChange={(e) => setTaxAmount(e.target.value)} className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500 mb-1 block">메모</label>
                <input type="text" value={memo} onChange={(e) => setMemo(e.target.value)} className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div className="flex items-end">
                <Button className="w-full" onClick={() => createInvoice.mutate()} disabled={!counterparty || createInvoice.isPending}>
                  {createInvoice.isPending ? '등록 중...' : '등록'}
                </Button>
              </div>
            </div>
          </Card>
        )}

        {/* 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-6 text-center text-sm text-neutral-400">로딩중...</div>
          ) : !data?.invoices.length ? (
            <div className="p-6 text-center text-sm text-neutral-400">세금계산서가 없습니다</div>
          ) : (
            <table className="w-full text-sm">
              <thead className="bg-neutral-50">
                <tr className="text-xs text-neutral-500">
                  <th className="text-left px-4 py-2 font-medium">유형</th>
                  <th className="text-left px-4 py-2 font-medium">발행일</th>
                  <th className="text-left px-4 py-2 font-medium">거래처</th>
                  <th className="text-right px-4 py-2 font-medium">공급가</th>
                  <th className="text-right px-4 py-2 font-medium">세액</th>
                  <th className="text-right px-4 py-2 font-medium">합계</th>
                  <th className="w-10"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-neutral-100">
                {data.invoices.map((inv) => (
                  <tr key={inv.id} className="hover:bg-warm-ivory/40">
                    <td className="px-4 py-2.5">
                      <span className={`px-1.5 py-0.5 rounded text-[10px] font-medium ${inv.invoice_type === 'sales' ? 'bg-blue-100 text-blue-700' : 'bg-orange-100 text-orange-700'}`}>
                        {inv.invoice_type === 'sales' ? '매출' : '매입'}
                      </span>
                    </td>
                    <td className="px-4 py-2.5 text-neutral-600">{formatDate(inv.issue_date)}</td>
                    <td className="px-4 py-2.5 font-medium">{inv.counterparty_name}</td>
                    <td className="px-4 py-2.5 text-right text-neutral-600">{formatKRW(inv.supply_amount)}</td>
                    <td className="px-4 py-2.5 text-right text-neutral-600">{formatKRW(inv.tax_amount)}</td>
                    <td className="px-4 py-2.5 text-right font-bold">{formatKRW(inv.total_amount)}</td>
                    <td className="px-2 py-2.5">
                      <button onClick={() => { if (confirm('삭제?')) deleteInvoice.mutate(inv.id); }}
                        className="w-6 h-6 rounded flex items-center justify-center text-neutral-400 hover:text-red-500">
                        <Trash2 size={12} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </Card>
      </div>
    </>
  );
}
