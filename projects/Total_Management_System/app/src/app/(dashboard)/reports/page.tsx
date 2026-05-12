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
  const [revenueTab, setRevenueTab] = useState('all');

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
            {/* 매출 탭 바 */}
            <div className="flex gap-1">
              {[
                { key: 'all', label: '전체 매출' },
                { key: 'product', label: '상품 판매' },
                { key: 'repair', label: '복원수리' },
              ].map((t) => (
                <button
                  key={t.key}
                  onClick={() => setRevenueTab(t.key)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition ${
                    revenueTab === t.key ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
                  }`}
                >
                  {t.label}
                </button>
              ))}
            </div>

            {/* 요약 카드 4개 */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
              {/* 매출 — 2단계 3분할: 제품(B2C+B2B) + 복원수리(접수+판매RS+납품RS) */}
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-lg bg-green-50 flex items-center justify-center">
                    <TrendingUp size={18} className="text-green-600" />
                  </div>
                  <h3 className="text-sm font-bold text-indigo-black">
                    {revenueTab === 'all' ? '전체 매출' : revenueTab === 'repair' ? '복원수리 매출' : '제품 매출'}
                  </h3>
                </div>
                {(() => {
                  const ps = data.product_sales;
                  const rep = data.repair_sales;
                  const productTotal = ps?.total ?? data.sales.total;
                  const repairTotal = rep?.total ?? 0;
                  const big = revenueTab === 'all' ? (data.total_revenue ?? (productTotal + repairTotal))
                    : revenueTab === 'repair' ? repairTotal
                    : productTotal;
                  return (
                    <>
                      <p className="text-2xl font-bold text-green-600">{formatKRW(big)}</p>
                      <div className="mt-2 space-y-1 text-xs text-neutral-500">
                        {revenueTab === 'all' && (
                          <>
                            <div className="flex justify-between font-medium text-neutral-600"><span>제품 매출</span><span>{formatKRW(productTotal)}</span></div>
                            {ps && (
                              <>
                                <div className="flex justify-between pl-2"><span>· B2C</span><span>{formatKRW(ps.b2c)}</span></div>
                                <div className="flex justify-between pl-2"><span>· B2B (납품 포함)</span><span>{formatKRW(ps.b2b)}</span></div>
                              </>
                            )}
                            <div className="flex justify-between font-medium text-neutral-600 pt-1"><span>복원수리</span><span>{formatKRW(repairTotal)}</span></div>
                            {rep && (
                              <>
                                <div className="flex justify-between pl-2"><span>· 접수시스템</span><span>{formatKRW(rep.a_intake)}</span></div>
                                <div className="flex justify-between pl-2"><span>· 판매 RS</span><span>{formatKRW(rep.b_offline_rs)}</span></div>
                                <div className="flex justify-between pl-2"><span>· 납품 RS</span><span>{formatKRW(rep.c_delivery_rs)}</span></div>
                              </>
                            )}
                          </>
                        )}
                        {revenueTab === 'product' && (
                          <>
                            {ps ? (
                              <>
                                <div className="flex justify-between"><span>B2C (소매·온라인)</span><span>{ps.offline_count}건 · {formatKRW(ps.b2c)}</span></div>
                                <div className="flex justify-between"><span>B2B (딜러·아카데미)</span><span>{formatKRW(ps.b2b_offline)}</span></div>
                                <div className="flex justify-between"><span>B2B (납품)</span><span>{ps.delivery_count}건 · {formatKRW(ps.b2b_delivery)}</span></div>
                              </>
                            ) : (
                              <div className="flex justify-between"><span>오프라인 판매</span><span>{data.sales.count}건 · {formatKRW(data.sales.total)}</span></div>
                            )}
                            <div className="flex justify-between pt-1 border-t border-neutral-100"><span>공급가액 <span className="text-neutral-300">(오프라인)</span></span><span>{formatKRW(data.sales.supply)}</span></div>
                            <div className="flex justify-between"><span>부가세 <span className="text-neutral-300">(오프라인)</span></span><span>{formatKRW(data.sales.vat)}</span></div>
                            <div className="flex justify-between"><span>할인</span><span>-{formatKRW(data.sales.discount)}</span></div>
                          </>
                        )}
                        {revenueTab === 'repair' && rep && (
                          <>
                            <div className="flex justify-between"><span>접수시스템</span><span>{rep.a_count}건 · {formatKRW(rep.a_intake)}</span></div>
                            <div className="flex justify-between"><span>판매 RS</span><span>{rep.b_count}건 · {formatKRW(rep.b_offline_rs)}</span></div>
                            <div className="flex justify-between"><span>납품 RS</span><span>{rep.c_count}건 · {formatKRW(rep.c_delivery_rs)}</span></div>
                            <div className="flex justify-between pt-1 border-t border-neutral-100"><span>수리비 <span className="text-neutral-300">(접수분)</span></span><span>{formatKRW(rep.service_cost_total)}</span></div>
                            <div className="flex justify-between"><span>수거·배송비 <span className="text-neutral-300">(접수분)</span></span><span>{formatKRW(rep.shipping_fee_total)}</span></div>
                          </>
                        )}
                      </div>
                    </>
                  );
                })()}
                {/* 결제방식별 — 제품/전체 탭 (오프라인 판매 기준) */}
                {revenueTab !== 'repair' && Object.keys(data.sales.by_method).length > 0 && (
                  <div className="mt-3 pt-3 border-t border-neutral-100">
                    <p className="text-[10px] text-neutral-400 mb-1">결제방식별 <span className="text-neutral-300">(오프라인 판매)</span></p>
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

              {/* 손익계산서 */}
              <Card>
                <div className="flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-lg bg-amber-50 flex items-center justify-center">
                    <BarChart3 size={18} className="text-amber-600" />
                  </div>
                  <h3 className="text-sm font-bold text-indigo-black">손익</h3>
                </div>
                {(() => {
                  const pl = (data as unknown as { profit_loss?: { revenue: number; cogs: number; gross_profit: number; expenses: number; operating_profit: number; margin_rate: number } }).profit_loss;
                  const opProfit = pl?.operating_profit ?? data.margin.gross_profit;
                  return (
                    <>
                      <p className={`text-2xl font-bold ${opProfit >= 0 ? 'text-amber-600' : 'text-red-500'}`}>
                        {formatKRW(opProfit)}
                      </p>
                      <p className="text-[10px] text-neutral-400 mt-0.5">
                        영업이익률 {pl?.margin_rate ?? data.margin.margin_rate}%
                      </p>
                      <div className="mt-2 space-y-1 text-xs text-neutral-500">
                        <div className="flex justify-between"><span>매출</span><span>{formatKRW(pl?.revenue ?? data.sales.total)}</span></div>
                        <div className="flex justify-between"><span>매출원가</span><span>-{formatKRW(pl?.cogs ?? data.margin.total_cogs)}</span></div>
                        <div className="flex justify-between font-semibold text-indigo-black border-t border-neutral-100 pt-1">
                          <span>매출총이익</span><span>{formatKRW(pl?.gross_profit ?? data.margin.gross_profit)}</span>
                        </div>
                        {pl && pl.expenses > 0 && (
                          <>
                            <div className="flex justify-between"><span>경비</span><span>-{formatKRW(pl.expenses)}</span></div>
                            <div className="flex justify-between font-bold text-indigo-black border-t border-neutral-200 pt-1">
                              <span>영업이익</span><span className={opProfit >= 0 ? '' : 'text-red-500'}>{formatKRW(opProfit)}</span>
                            </div>
                          </>
                        )}
                      </div>
                    </>
                  );
                })()}
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

                {/* 미수금 + 에이징 */}
                <Card>
                  <div className="flex items-center gap-2 mb-3">
                    <Users size={16} className="text-amber-600" />
                    <h3 className="text-sm font-bold text-indigo-black">미수금</h3>
                    <span className="ml-auto text-sm font-bold text-amber-600">{formatKRW(data.receivables.total)}</span>
                  </div>
                  {/* 에이징 바 */}
                  {(() => {
                    const aging = (data.receivables as unknown as { aging?: { within30: number; d30to60: number; d60to90: number; over90: number } }).aging;
                    if (!aging) return null;
                    return (
                      <div className="grid grid-cols-4 gap-1 mb-3">
                        {[
                          { label: '30일이내', amount: aging.within30, color: 'bg-green-100 text-green-700' },
                          { label: '30~60일', amount: aging.d30to60, color: 'bg-yellow-100 text-yellow-700' },
                          { label: '60~90일', amount: aging.d60to90, color: 'bg-orange-100 text-orange-700' },
                          { label: '90일+', amount: aging.over90, color: 'bg-red-100 text-red-700' },
                        ].map((a) => (
                          <div key={a.label} className={`rounded p-1.5 text-center ${a.color}`}>
                            <p className="text-[10px] font-medium">{a.label}</p>
                            <p className="text-xs font-bold">{formatKRW(a.amount || 0)}</p>
                          </div>
                        ))}
                      </div>
                    );
                  })()}
                  {data.receivables.items.length === 0 ? (
                    <p className="text-xs text-neutral-400 text-center py-3">미수금 없음</p>
                  ) : (
                    <div className="space-y-2 max-h-48 overflow-y-auto">
                      {data.receivables.items.map((r) => {
                        const days = (r as unknown as Record<string, unknown>).daysOverdue as number || 0;
                        const agingColor = days <= 30 ? 'text-green-600' : days <= 60 ? 'text-yellow-600' : days <= 90 ? 'text-orange-600' : 'text-red-600';
                        return (
                          <Link key={r.id} href={`/customers/${r.id}`} className="flex items-center justify-between py-1.5 border-b border-neutral-50 last:border-0 hover:bg-warm-ivory/40 transition rounded px-1">
                            <div>
                              <p className="text-sm font-medium">{r.name}</p>
                              {days > 0 && <p className={`text-[10px] font-medium ${agingColor}`}>{days}일 경과</p>}
                            </div>
                            <span className="text-sm font-bold text-amber-600">{formatKRW(r.outstanding)}</span>
                          </Link>
                        );
                      })}
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
