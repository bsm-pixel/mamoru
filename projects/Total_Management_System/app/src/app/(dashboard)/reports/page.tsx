'use client';

import { useState } from 'react';
import Link from 'next/link';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useReportSummary, downloadExcel } from '@/hooks/use-reports';
import { formatKRW } from '@/lib/utils/format';
import { Download, FileSpreadsheet, TrendingUp, TrendingDown, Receipt, FileText, BarChart3, Truck, Wallet, Users } from 'lucide-react';

const METHOD_LABEL: Record<string, string> = {
  card: '카드', cash: '현금', transfer: '계좌이체', mixed: '복합',
};

// 기간 프리셋
function getPreset(key: string): { from: string; to: string } {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();

  switch (key) {
    case 'this_month':
      return {
        from: new Date(y, m, 1).toISOString().slice(0, 10),
        to: now.toISOString().slice(0, 10),
      };
    case 'last_month':
      return {
        from: new Date(y, m - 1, 1).toISOString().slice(0, 10),
        to: new Date(y, m, 0).toISOString().slice(0, 10),
      };
    case 'this_quarter': {
      const qStart = Math.floor(m / 3) * 3;
      return {
        from: new Date(y, qStart, 1).toISOString().slice(0, 10),
        to: now.toISOString().slice(0, 10),
      };
    }
    case 'this_year':
      return {
        from: new Date(y, 0, 1).toISOString().slice(0, 10),
        to: now.toISOString().slice(0, 10),
      };
    default:
      return {
        from: new Date(y, m, 1).toISOString().slice(0, 10),
        to: now.toISOString().slice(0, 10),
      };
  }
}

const PRESETS = [
  { key: 'this_month', label: '이번 달' },
  { key: 'last_month', label: '지난 달' },
  { key: 'this_quarter', label: '이번 분기' },
  { key: 'this_year', label: '올해' },
];

export default function ReportsPage() {
  const [preset, setPreset] = useState('this_month');
  const [customFrom, setCustomFrom] = useState('');
  const [customTo, setCustomTo] = useState('');

  const period = customFrom && customTo
    ? { from: customFrom, to: customTo }
    : getPreset(preset);

  const { data, isLoading } = useReportSummary(period.from, period.to);

  return (
    <>
      <Topbar title="회계" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 기간 선택 */}
        <Card>
          <div className="flex flex-wrap items-center gap-2">
            {PRESETS.map((p) => (
              <button
                key={p.key}
                onClick={() => { setPreset(p.key); setCustomFrom(''); setCustomTo(''); }}
                className={`px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                  preset === p.key && !customFrom
                    ? 'bg-terracotta text-cream'
                    : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                }`}
              >
                {p.label}
              </button>
            ))}
            <div className="flex items-center gap-1.5 ml-auto">
              <input
                type="date"
                value={customFrom || period.from}
                onChange={(e) => { setCustomFrom(e.target.value); setPreset(''); }}
                className="h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
              <span className="text-xs text-neutral-400">~</span>
              <input
                type="date"
                value={customTo || period.to}
                onChange={(e) => { setCustomTo(e.target.value); setPreset(''); }}
                className="h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
          </div>
        </Card>

        {isLoading ? (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            {Array.from({ length: 4 }).map((_, i) => <Skeleton key={i} className="h-36" />)}
          </div>
        ) : data ? (
          <>
            {/* 요약 카드 4개 */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
              {/* 매출 */}
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-lg bg-green-50 flex items-center justify-center">
                    <TrendingUp size={18} className="text-green-600" />
                  </div>
                  <h3 className="text-sm font-bold text-indigo-black">매출</h3>
                </div>
                <p className="text-2xl font-bold text-green-600">{formatKRW(data.sales.total)}</p>
                <div className="mt-2 space-y-1 text-xs text-neutral-500">
                  <div className="flex justify-between"><span>판매 건수</span><span>{data.sales.count}건</span></div>
                  <div className="flex justify-between"><span>공급가액</span><span>{formatKRW(data.sales.supply)}</span></div>
                  <div className="flex justify-between"><span>부가세</span><span>{formatKRW(data.sales.vat)}</span></div>
                  <div className="flex justify-between"><span>할인</span><span>-{formatKRW(data.sales.discount)}</span></div>
                </div>
                {/* 결제방식별 */}
                {Object.keys(data.sales.by_method).length > 0 && (
                  <div className="mt-3 pt-3 border-t border-neutral-100">
                    <p className="text-[10px] text-neutral-400 mb-1">결제방식별</p>
                    <div className="space-y-0.5">
                      {Object.entries(data.sales.by_method).map(([k, v]) => (
                        <div key={k} className="flex justify-between text-xs">
                          <span className="text-neutral-500">{METHOD_LABEL[k] || k}</span>
                          <span className="font-medium">{formatKRW(v)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </Card>

              {/* 매입 */}
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-lg bg-red-50 flex items-center justify-center">
                    <TrendingDown size={18} className="text-red-500" />
                  </div>
                  <h3 className="text-sm font-bold text-indigo-black">매입</h3>
                </div>
                <p className="text-2xl font-bold text-red-500">{formatKRW(data.purchases.total)}</p>
                <div className="mt-2 space-y-1 text-xs text-neutral-500">
                  <div className="flex justify-between"><span>발주 건수</span><span>{data.purchases.count}건</span></div>
                  <div className="flex justify-between"><span>공급가액</span><span>{formatKRW(data.purchases.supply)}</span></div>
                  <div className="flex justify-between"><span>부가세</span><span>{formatKRW(data.purchases.vat)}</span></div>
                  <div className="flex justify-between"><span>선납금</span><span>{formatKRW(data.purchases.deposit)}</span></div>
                </div>
              </Card>

              {/* VAT */}
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-lg bg-indigo-50 flex items-center justify-center">
                    <Receipt size={18} className="text-indigo-600" />
                  </div>
                  <h3 className="text-sm font-bold text-indigo-black">부가세 요약</h3>
                </div>
                <p className={`text-2xl font-bold ${data.vat.net_vat >= 0 ? 'text-indigo-600' : 'text-green-600'}`}>
                  {data.vat.net_vat >= 0 ? '' : '-'}{formatKRW(Math.abs(data.vat.net_vat))}
                </p>
                <p className="text-[10px] text-neutral-400 mt-0.5">
                  {data.vat.net_vat >= 0 ? '납부 예상' : '환급 예상'}
                </p>
                <div className="mt-2 space-y-1 text-xs text-neutral-500">
                  <div className="flex justify-between"><span>매출세액</span><span>{formatKRW(data.vat.sales_vat)}</span></div>
                  <div className="flex justify-between"><span>매입세액</span><span>-{formatKRW(data.vat.purchase_vat)}</span></div>
                </div>

              </Card>

              {/* 매출이익 (COGS 기반) */}
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-lg bg-amber-50 flex items-center justify-center">
                    <BarChart3 size={18} className="text-amber-600" />
                  </div>
                  <h3 className="text-sm font-bold text-indigo-black">매출이익</h3>
                </div>
                <p className={`text-2xl font-bold ${data.margin.gross_profit >= 0 ? 'text-amber-600' : 'text-red-500'}`}>
                  {formatKRW(data.margin.gross_profit)}
                </p>
                <p className="text-[10px] text-neutral-400 mt-0.5">
                  이익률 {data.margin.margin_rate}%
                </p>
                <div className="mt-2 space-y-1 text-xs text-neutral-500">
                  <div className="flex justify-between"><span>매출</span><span>{formatKRW(data.sales.total)}</span></div>
                  <div className="flex justify-between"><span>매출원가 (COGS)</span><span>-{formatKRW(data.margin.total_cogs)}</span></div>
                  <div className="flex justify-between font-semibold text-indigo-black">
                    <span>매출총이익</span>
                    <span>{formatKRW(data.margin.gross_profit)}</span>
                  </div>
                </div>
              </Card>
            </div>

            {/* 일별 추이 (간단 바 차트) */}
            <Card>
              <h3 className="text-sm font-bold text-indigo-black mb-3">일별 매출/매입 추이</h3>
              <DailyChart sales={data.daily.sales} purchases={data.daily.purchases} />
            </Card>

            {/* 엑셀 다운로드 + 거래내역서 */}
            <div className="flex flex-wrap gap-2">
              <Button
                variant="secondary"
                size="sm"
                onClick={() => downloadExcel('sales', period.from, period.to)}
              >
                <FileSpreadsheet size={14} />
                매출 엑셀
                <Download size={12} />
              </Button>
              <Button
                variant="secondary"
                size="sm"
                onClick={() => downloadExcel('purchases', period.from, period.to)}
              >
                <FileSpreadsheet size={14} />
                매입 엑셀
                <Download size={12} />
              </Button>
              <Button
                variant="secondary"
                size="sm"
                onClick={() => downloadExcel('margin' as 'sales', period.from, period.to)}
              >
                <BarChart3 size={14} />
                마진 분석
                <Download size={12} />
              </Button>
              <Link href="/reports/transaction">
                <Button variant="secondary" size="sm">
                  <FileText size={14} />
                  거래내역서
                </Button>
              </Link>
            </div>

            {/* 제품별 매출 랭킹 */}
            {data.by_product.length > 0 && (
              <Card>
                <h3 className="text-sm font-bold text-indigo-black mb-3">제품별 매출 랭킹</h3>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="border-b border-neutral-200 text-neutral-500">
                        <th className="text-left py-2 pr-2">제품</th>
                        <th className="text-right py-2 px-2">수량</th>
                        <th className="text-right py-2 px-2">매출</th>
                        <th className="text-right py-2 px-2">원가</th>
                        <th className="text-right py-2 px-2">이익</th>
                        <th className="text-right py-2 pl-2">이익률</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-neutral-50">
                      {data.by_product.slice(0, 10).map((p) => (
                        <tr key={p.product_id} className="hover:bg-warm-ivory/40">
                          <td className="py-2 pr-2">
                            <p className="font-medium text-indigo-black truncate max-w-[140px]">{p.product_name}</p>
                            <p className="text-[10px] text-neutral-400">{p.sku}</p>
                          </td>
                          <td className="text-right py-2 px-2">{p.qty}</td>
                          <td className="text-right py-2 px-2 font-semibold">{formatKRW(p.revenue)}</td>
                          <td className="text-right py-2 px-2 text-neutral-500">{formatKRW(p.cogs)}</td>
                          <td className={`text-right py-2 px-2 font-semibold ${p.profit >= 0 ? 'text-green-600' : 'text-red-500'}`}>
                            {formatKRW(p.profit)}
                          </td>
                          <td className={`text-right py-2 pl-2 ${p.margin_rate >= 30 ? 'text-green-600' : p.margin_rate >= 0 ? 'text-amber-600' : 'text-red-500'}`}>
                            {p.margin_rate}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {/* 매입처별 지출 */}
            {data.by_supplier.length > 0 && (
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <Truck size={16} className="text-neutral-500" />
                  <h3 className="text-sm font-bold text-indigo-black">매입처별 지출</h3>
                </div>
                <div className="space-y-2">
                  {data.by_supplier.map((s) => (
                    <div key={s.name} className="flex items-center justify-between py-1.5 border-b border-neutral-50 last:border-0">
                      <div>
                        <p className="text-sm font-medium">{s.name}</p>
                        <p className="text-[10px] text-neutral-400">{s.count}건</p>
                      </div>
                      <span className="text-sm font-bold text-red-500">{formatKRW(s.total)}</span>
                    </div>
                  ))}
                </div>
              </Card>
            )}

            {/* 미지급금 / 미수금 */}
            {(data.payables.total > 0 || data.receivables.total > 0) && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* 미지급금 */}
                <Card>
                  <div className="flex items-center gap-2 mb-3">
                    <Wallet size={16} className="text-red-500" />
                    <h3 className="text-sm font-bold text-indigo-black">미지급금</h3>
                    <span className="ml-auto text-sm font-bold text-red-500">{formatKRW(data.payables.total)}</span>
                  </div>
                  {data.payables.items.length === 0 ? (
                    <p className="text-xs text-neutral-400 text-center py-3">미지급금 없음</p>
                  ) : (
                    <div className="space-y-2 max-h-48 overflow-y-auto">
                      {data.payables.items.map((p) => (
                        <div key={p.name} className="flex items-center justify-between py-1.5 border-b border-neutral-50 last:border-0">
                          <div>
                            <p className="text-sm font-medium">{p.name}</p>
                            <p className="text-[10px] text-neutral-400">{p.count}건 미완료</p>
                          </div>
                          <span className="text-sm font-bold text-red-500">{formatKRW(p.total_owed)}</span>
                        </div>
                      ))}
                    </div>
                  )}
                </Card>

                {/* 미수금 */}
                <Card>
                  <div className="flex items-center gap-2 mb-3">
                    <Users size={16} className="text-amber-600" />
                    <h3 className="text-sm font-bold text-indigo-black">미수금</h3>
                    <span className="ml-auto text-sm font-bold text-amber-600">{formatKRW(data.receivables.total)}</span>
                  </div>
                  {data.receivables.items.length === 0 ? (
                    <p className="text-xs text-neutral-400 text-center py-3">미수금 없음</p>
                  ) : (
                    <div className="space-y-2 max-h-48 overflow-y-auto">
                      {data.receivables.items.map((r) => (
                        <Link key={r.id} href={`/customers/${r.id}`} className="flex items-center justify-between py-1.5 border-b border-neutral-50 last:border-0 hover:bg-warm-ivory/40 transition rounded px-1">
                          <p className="text-sm font-medium">{r.name}</p>
                          <span className="text-sm font-bold text-amber-600">{formatKRW(r.outstanding)}</span>
                        </Link>
                      ))}
                    </div>
                  )}
                </Card>
              </div>
            )}

            {/* 최근 매출 내역 */}
            <Card>
              <h3 className="text-sm font-bold text-indigo-black mb-3">
                매출 내역 ({data.details.sales.length}건)
              </h3>
              {data.details.sales.length === 0 ? (
                <p className="text-xs text-neutral-400 text-center py-4">해당 기간 매출 내역이 없습니다</p>
              ) : (
                <div className="space-y-2 max-h-64 overflow-y-auto">
                  {data.details.sales.map((s) => (
                    <Link key={s.id} href={`/sales/${s.id}`} className="flex items-center justify-between py-1.5 border-b border-neutral-50 last:border-0 hover:bg-warm-ivory/40 transition rounded px-1">
                      <div>
                        <p className="text-sm font-medium">{s.customer_name}</p>
                        <p className="text-xs text-neutral-500">{s.sale_date}</p>
                      </div>
                      <span className="text-sm font-bold">{formatKRW(s.total_amount)}</span>
                    </Link>
                  ))}
                </div>
              )}
            </Card>

            {/* 최근 매입 내역 */}
            <Card>
              <h3 className="text-sm font-bold text-indigo-black mb-3">
                매입 내역 ({data.details.purchases.length}건)
              </h3>
              {data.details.purchases.length === 0 ? (
                <p className="text-xs text-neutral-400 text-center py-4">해당 기간 매입 내역이 없습니다</p>
              ) : (
                <div className="space-y-2 max-h-64 overflow-y-auto">
                  {data.details.purchases.map((p) => (
                    <Link key={p.id} href={`/purchasing/${p.id}`} className="flex items-center justify-between py-1.5 border-b border-neutral-50 last:border-0 hover:bg-warm-ivory/40 transition rounded px-1">
                      <div>
                        <p className="text-sm font-medium">{p.supplier_name}</p>
                        <p className="text-xs text-neutral-500">{p.order_date}</p>
                      </div>
                      <span className="text-sm font-bold">{formatKRW(p.total_amount)}</span>
                    </Link>
                  ))}
                </div>
              )}
            </Card>
          </>
        ) : null}
      </div>
    </>
  );
}

/* ---------- 일별 차트 ---------- */

function DailyChart({ sales, purchases }: { sales: Record<string, number>; purchases: Record<string, number> }) {
  const allDates = [...new Set([...Object.keys(sales), ...Object.keys(purchases)])].sort();

  if (allDates.length === 0) {
    return <p className="text-xs text-neutral-400 text-center py-4">데이터 없음</p>;
  }

  const maxVal = Math.max(
    ...allDates.map((d) => Math.max(sales[d] || 0, purchases[d] || 0)),
    1
  );

  return (
    <div className="space-y-1.5 max-h-60 overflow-y-auto">
      {allDates.map((date) => {
        const saleVal = sales[date] || 0;
        const purchVal = purchases[date] || 0;
        const salePct = (saleVal / maxVal) * 100;
        const purchPct = (purchVal / maxVal) * 100;

        return (
          <div key={date} className="flex items-center gap-2 text-xs">
            <span className="w-14 text-neutral-500 shrink-0">{date.slice(5)}</span>
            <div className="flex-1 space-y-0.5">
              {saleVal > 0 && (
                <div className="flex items-center gap-1">
                  <div
                    className="h-3 rounded bg-green-400/70"
                    style={{ width: `${Math.max(salePct, 2)}%` }}
                  />
                  <span className="text-[10px] text-green-600 shrink-0">{formatKRW(saleVal)}</span>
                </div>
              )}
              {purchVal > 0 && (
                <div className="flex items-center gap-1">
                  <div
                    className="h-3 rounded bg-red-400/70"
                    style={{ width: `${Math.max(purchPct, 2)}%` }}
                  />
                  <span className="text-[10px] text-red-500 shrink-0">{formatKRW(purchVal)}</span>
                </div>
              )}
              {saleVal === 0 && purchVal === 0 && (
                <span className="text-neutral-300">-</span>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}
