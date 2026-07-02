'use client';

import { useState, useMemo } from 'react';
import Link from 'next/link';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Skeleton } from '@/components/ui/skeleton';
import { StatCard } from '@/components/ui/stat-card';
import { useReportSummary, type ProductMargin } from '@/hooks/use-reports';
import { formatKRW, toLocalDateString } from '@/lib/utils/format';
import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';
import { Package, Boxes, TrendingUp, Trophy, ArrowLeft, Search } from 'lucide-react';

// 기간 프리셋 — toLocalDateString(KST). reports/page.tsx 와 동일 로직
function getPreset(key: string): { from: string; to: string } {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  const today = toLocalDateString(now);
  switch (key) {
    case 'this_month': return { from: toLocalDateString(new Date(y, m, 1)), to: today };
    case 'last_month': return { from: toLocalDateString(new Date(y, m - 1, 1)), to: toLocalDateString(new Date(y, m, 0)) };
    case 'this_quarter': { const q = Math.floor(m / 3) * 3; return { from: toLocalDateString(new Date(y, q, 1)), to: today }; }
    case 'this_year': return { from: toLocalDateString(new Date(y, 0, 1)), to: today };
    default: return { from: toLocalDateString(new Date(y, m, 1)), to: today };
  }
}
const PRESETS = [
  { key: 'this_month', label: '이번 달' },
  { key: 'last_month', label: '지난 달' },
  { key: 'this_quarter', label: '이번 분기' },
  { key: 'this_year', label: '올해' },
];
const SORTS = [
  { key: 'qty', label: '인기순(수량)' },
  { key: 'revenue', label: '매출순' },
  { key: 'profit', label: '이익순' },
  { key: 'margin', label: '이익률순' },
];
const CHANNELS = [
  { key: 'all', label: '전체' },
  { key: 'b2c', label: 'B2C 소매·온라인' },
  { key: 'b2b', label: 'B2B 딜러·납품' },
];
// 선택 채널 기준 수치 (전체/B2C/B2B)
function viewOf(p: ProductMargin, ch: string) {
  const qty = ch === 'b2c' ? (p.b2c_qty || 0) : ch === 'b2b' ? (p.b2b_qty || 0) : p.qty;
  const revenue = ch === 'b2c' ? (p.b2c_revenue || 0) : ch === 'b2b' ? (p.b2b_revenue || 0) : p.revenue;
  const cogs = ch === 'b2c' ? (p.b2c_cogs || 0) : ch === 'b2b' ? (p.b2b_cogs || 0) : p.cogs;
  const profit = revenue - cogs;
  const margin_rate = revenue > 0 ? Math.round((profit / revenue) * 1000) / 10 : 0;
  return { qty, revenue, cogs, profit, margin_rate };
}

export default function ProductSalesPage() {
  const [preset, setPreset] = useState('this_month');
  const [customFrom, setCustomFrom] = useState('');
  const [customTo, setCustomTo] = useState('');
  const [q, setQ] = useState('');
  const [sort, setSort] = useState('qty');
  const [cat, setCat] = useState('all');
  const [channel, setChannel] = useState('all');

  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const period = customFrom && customTo ? { from: customFrom, to: customTo } : getPreset(preset);
  const { data, isLoading } = useReportSummary(period.from, period.to);

  const totals = useMemo(() => {
    const bp = data?.by_product || [];
    let qty = 0, revenue = 0, count = 0, topQ = -1;
    let top: { product_name: string; qty: number; revenue: number } | null = null;
    for (const p of bp) {
      const v = viewOf(p, channel);
      if (v.qty > 0 || v.revenue > 0) count++;
      qty += v.qty; revenue += v.revenue;
      if (v.qty > topQ) { topQ = v.qty; top = { product_name: p.product_name, qty: v.qty, revenue: v.revenue }; }
    }
    return { count, qty, revenue, top };
  }, [data, channel]);

  const cats = useMemo(() => {
    const set = new Set<string>();
    (data?.by_product || []).forEach((p) => { if (p.category) set.add(p.category); });
    return [...set];
  }, [data]);

  const rows = useMemo(() => {
    let list = data?.by_product ? [...data.by_product] : [];
    if (cat !== 'all') list = list.filter((p) => (p.category || '') === cat);
    const s = q.trim().toLowerCase();
    if (s) list = list.filter((p) => (p.product_name || '').toLowerCase().includes(s) || (p.sku || '').toLowerCase().includes(s));
    if (channel !== 'all') list = list.filter((p) => { const v = viewOf(p, channel); return v.qty > 0 || v.revenue > 0; });
    list.sort((a, b) => {
      const va = viewOf(a, channel), vb = viewOf(b, channel);
      return sort === 'revenue' ? vb.revenue - va.revenue
        : sort === 'profit' ? vb.profit - va.profit
        : sort === 'margin' ? vb.margin_rate - va.margin_rate
        : vb.qty - va.qty;
    });
    return list;
  }, [data, cat, q, sort, channel]);

  return (
    <>
      <Topbar title="품목별 매출" />
      <div className="px-4 md:px-6 py-4 space-y-4">
        <Link href="/reports" className="inline-flex items-center gap-1 text-xs text-neutral-500 hover:text-neutral-900">
          <ArrowLeft size={14} /> 회계로
        </Link>

        {/* 기간 선택 */}
        <Card>
          <div className="flex flex-wrap items-center gap-2">
            {PRESETS.map((p) => (
              <button key={p.key} onClick={() => { setPreset(p.key); setCustomFrom(''); setCustomTo(''); }}
                className={`px-3 py-1.5 rounded-full text-xs font-semibold transition ${preset === p.key && !customFrom ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                {p.label}
              </button>
            ))}
            <div className="flex items-center gap-1.5 ml-auto">
              <input type="date" value={customFrom || period.from} onChange={(e) => { setCustomFrom(e.target.value); setPreset(''); }}
                className="h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-stone-400" />
              <span className="text-xs text-neutral-400">~</span>
              <input type="date" value={customTo || period.to} onChange={(e) => { setCustomTo(e.target.value); setPreset(''); }}
                className="h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-xs focus:outline-none focus:ring-2 focus:ring-stone-400" />
            </div>
          </div>
        </Card>

        {isLoading ? (
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
            {Array.from({ length: 4 }).map((_, i) => <Skeleton key={i} className="h-28" />)}
          </div>
        ) : !data ? (
          <p className="text-sm text-neutral-400 text-center py-16">데이터를 불러오지 못했습니다</p>
        ) : (
          <>
            {/* 요약 카드 */}
            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
              <StatCard label="판매 품목수" icon={Package} accent="blue" value={totals.count} primarySub="종류" />
              <StatCard label="총 판매수량" icon={Boxes} accent="violet" value={totals.qty} primarySub="개" />
              <StatCard label="제품 매출" icon={TrendingUp} accent="emerald" value={formatKRW(totals.revenue)} primarySub="RS 제외" />
              <StatCard label="최다 판매" icon={Trophy} accent="amber"
                value={totals.top ? totals.top.product_name : '-'}
                primarySub={totals.top ? `${totals.top.qty}개 · ${formatKRW(totals.top.revenue)}` : ''} />
            </div>

            {/* 검색 + 정렬 + 카테고리 */}
            <Card>
              <div className="flex flex-col gap-3">
                <div className="flex items-center gap-2 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50">
                  <Search size={15} className="text-neutral-400 shrink-0" />
                  <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="제품명 · SKU 검색"
                    className="flex-1 bg-transparent text-sm outline-none placeholder:text-neutral-400" />
                </div>
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="text-[11px] text-neutral-400 mr-1">정렬</span>
                  {SORTS.map((s) => (
                    <button key={s.key} onClick={() => setSort(s.key)}
                      className={`px-2.5 py-1 rounded-full text-xs font-semibold transition ${sort === s.key ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                      {s.label}
                    </button>
                  ))}
                </div>
                <div className="flex flex-wrap items-center gap-1.5">
                  <span className="text-[11px] text-neutral-400 mr-1">채널</span>
                  {CHANNELS.map((c) => (
                    <button key={c.key} onClick={() => setChannel(c.key)}
                      className={`px-2.5 py-1 rounded-full text-xs font-semibold transition ${channel === c.key ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                      {c.label}
                    </button>
                  ))}
                </div>
                {cats.length > 0 && (
                  <div className="flex flex-wrap items-center gap-1.5">
                    <span className="text-[11px] text-neutral-400 mr-1">분류</span>
                    <button onClick={() => setCat('all')}
                      className={`px-2.5 py-1 rounded-full text-xs font-semibold transition ${cat === 'all' ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>전체</button>
                    {cats.map((c) => (
                      <button key={c} onClick={() => setCat(c)}
                        className={`px-2.5 py-1 rounded-full text-xs font-semibold transition ${cat === c ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                        {catLabels[c] || c}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </Card>

            {/* 전체 목록 */}
            <Card>
              <div className="flex items-center justify-between mb-3">
                <h3 className="text-sm font-bold text-stone-900">품목별 매출 ({rows.length})</h3>
                <span className="text-[11px] text-neutral-400">{period.from} ~ {period.to} · 온라인·매장·납품 합산 (RS 제외)</span>
              </div>
              {rows.length === 0 ? (
                <p className="text-sm text-neutral-400 text-center py-12">해당 조건의 판매 품목이 없습니다</p>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="border-b border-neutral-200 text-neutral-500">
                        <th className="text-left py-2 pr-2 w-8">#</th>
                        <th className="text-left py-2 pr-2">제품</th>
                        <th className="text-right py-2 px-2">수량</th>
                        <th className="text-right py-2 px-2">매출</th>
                        <th className="text-right py-2 px-2">원가</th>
                        <th className="text-right py-2 px-2">이익</th>
                        <th className="text-right py-2 pl-2">이익률</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-neutral-50">
                      {rows.map((p, i) => {
                        const v = viewOf(p, channel);
                        return (
                        <tr key={p.product_id + String(i)} className="hover:bg-stone-50">
                          <td className="py-2 pr-2 text-neutral-400 font-mono">{i + 1}</td>
                          <td className="py-2 pr-2">
                            <p className="font-medium text-stone-900 truncate max-w-[220px]">{p.product_name}</p>
                            <p className="text-[10px] text-neutral-400">{p.sku}{p.category ? ` · ${catLabels[p.category] || p.category}` : ''}</p>
                            {channel === 'all' && ((p.b2c_qty || 0) > 0 || (p.b2b_qty || 0) > 0) && (
                              <p className="text-[10px] mt-0.5">
                                <span className="text-blue-500">B2C {p.b2c_qty || 0}개·{formatKRW(p.b2c_revenue || 0)}</span>
                                <span className="text-neutral-300"> / </span>
                                <span className="text-violet-500">B2B {p.b2b_qty || 0}개·{formatKRW(p.b2b_revenue || 0)}</span>
                              </p>
                            )}
                          </td>
                          <td className="text-right py-2 px-2 font-semibold">{v.qty}</td>
                          <td className="text-right py-2 px-2 font-semibold">{formatKRW(v.revenue)}</td>
                          <td className="text-right py-2 px-2 text-neutral-500">{formatKRW(v.cogs)}</td>
                          <td className={`text-right py-2 px-2 font-semibold ${v.profit >= 0 ? 'text-emerald-600' : 'text-red-500'}`}>{formatKRW(v.profit)}</td>
                          <td className={`text-right py-2 pl-2 ${v.margin_rate >= 30 ? 'text-emerald-600' : v.margin_rate >= 0 ? 'text-amber-600' : 'text-red-500'}`}>{v.margin_rate}%</td>
                        </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </Card>
          </>
        )}
      </div>
    </>
  );
}
