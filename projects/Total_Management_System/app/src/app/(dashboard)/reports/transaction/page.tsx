'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useReportSummary, type SaleDetail } from '@/hooks/use-reports';
import { formatKRW } from '@/lib/utils/format';
import { ArrowLeft, Printer } from 'lucide-react';
import Link from 'next/link';

export default function TransactionPage() {
  const now = new Date();
  const [from, setFrom] = useState(new Date(now.getFullYear(), now.getMonth(), 1).toISOString().slice(0, 10));
  const [to, setTo] = useState(now.toISOString().slice(0, 10));
  const [customerFilter, setCustomerFilter] = useState('');

  const { data, isLoading } = useReportSummary(from, to);

  // 고객별 그룹핑
  const sales = data?.details.sales || [];
  const filtered = customerFilter
    ? sales.filter((s) => s.customer_name.includes(customerFilter))
    : sales;

  // 고객별 그룹
  const grouped = new Map<string, SaleDetail[]>();
  for (const s of filtered) {
    const key = s.customer_name || '미지정';
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key)!.push(s);
  }

  const grandTotal = filtered.reduce((s, r) => s + r.total_amount, 0);
  const grandSupply = filtered.reduce((s, r) => s + (r.supply_amount || 0), 0);
  const grandVat = filtered.reduce((s, r) => s + (r.vat_amount || 0), 0);

  return (
    <>
      <Topbar title="거래내역서" />

      <div className="px-4 md:px-6 py-4 space-y-4">
        {/* 뒤로가기 + 인쇄 */}
        <div className="flex items-center justify-between print:hidden">
          <Link href="/reports">
            <Button variant="ghost" size="sm">
              <ArrowLeft size={14} />
              리포트
            </Button>
          </Link>
          <Button size="sm" onClick={() => window.print()}>
            <Printer size={14} />
            인쇄
          </Button>
        </div>

        {/* 필터 (인쇄 시 숨김) */}
        <Card className="print:hidden">
          <div className="flex flex-wrap items-center gap-3">
            <div>
              <label className="text-xs text-neutral-500">시작일</label>
              <input
                type="date"
                value={from}
                onChange={(e) => setFrom(e.target.value)}
                className="block h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
            <div>
              <label className="text-xs text-neutral-500">종료일</label>
              <input
                type="date"
                value={to}
                onChange={(e) => setTo(e.target.value)}
                className="block h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
            <div className="flex-1">
              <label className="text-xs text-neutral-500">고객명 필터</label>
              <input
                type="text"
                value={customerFilter}
                onChange={(e) => setCustomerFilter(e.target.value)}
                placeholder="고객명 검색"
                className="block w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40"
              />
            </div>
          </div>
        </Card>

        {isLoading ? (
          <Skeleton className="h-96" />
        ) : (
          /* 인쇄 영역 */
          <div className="print-area bg-white">
            {/* 인쇄 헤더 (화면에서 숨김, 인쇄 시 표시) */}
            <div className="hidden print:block mb-6">
              <h1 className="text-xl font-bold text-center">거 래 내 역 서</h1>
              <div className="flex justify-between mt-4 text-sm">
                <div>
                  <p className="font-semibold">MAMORU (마모루)</p>
                  <p className="text-neutral-600">미용가위 전문</p>
                </div>
                <div className="text-right">
                  <p>기간: {from} ~ {to}</p>
                  <p>발행일: {new Date().toISOString().slice(0, 10)}</p>
                </div>
              </div>
            </div>

            {grouped.size === 0 ? (
              <p className="text-sm text-neutral-400 text-center py-8">해당 기간 거래 내역이 없습니다</p>
            ) : (
              <div className="space-y-6">
                {Array.from(grouped.entries()).map(([customerName, customerSales]) => {
                  const custTotal = customerSales.reduce((s, r) => s + r.total_amount, 0);
                  const custSupply = customerSales.reduce((s, r) => s + (r.supply_amount || 0), 0);
                  const custVat = customerSales.reduce((s, r) => s + (r.vat_amount || 0), 0);

                  return (
                    <div key={customerName} className="break-inside-avoid">
                      <h3 className="text-sm font-bold text-indigo-black mb-2 pb-1 border-b-2 border-indigo-black print:border-black">
                        {customerName}
                      </h3>
                      <table className="w-full text-xs">
                        <thead>
                          <tr className="border-b border-neutral-200 text-neutral-500">
                            <th className="text-left py-1.5 font-medium">일자</th>
                            <th className="text-right py-1.5 font-medium">합계</th>
                            <th className="text-right py-1.5 font-medium">공급가액</th>
                            <th className="text-right py-1.5 font-medium">부가세</th>
                            <th className="text-left py-1.5 font-medium pl-3">결제</th>
                          </tr>
                        </thead>
                        <tbody>
                          {customerSales.map((s) => (
                            <tr key={s.id} className="border-b border-neutral-50">
                              <td className="py-1.5">{s.sale_date}</td>
                              <td className="text-right py-1.5 font-medium">{formatKRW(s.total_amount)}</td>
                              <td className="text-right py-1.5 text-neutral-500">{formatKRW(s.supply_amount || 0)}</td>
                              <td className="text-right py-1.5 text-neutral-500">{formatKRW(s.vat_amount || 0)}</td>
                              <td className="py-1.5 pl-3 text-neutral-500">
                                {({ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' })[s.payment_method] || s.payment_method}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="border-t border-neutral-300 font-semibold">
                            <td className="py-1.5">소계 ({customerSales.length}건)</td>
                            <td className="text-right py-1.5">{formatKRW(custTotal)}</td>
                            <td className="text-right py-1.5 text-neutral-600">{formatKRW(custSupply)}</td>
                            <td className="text-right py-1.5 text-neutral-600">{formatKRW(custVat)}</td>
                            <td></td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  );
                })}

                {/* 총계 */}
                <div className="border-t-2 border-indigo-black print:border-black pt-3">
                  <table className="w-full text-sm font-bold">
                    <tbody>
                      <tr>
                        <td>총계 ({filtered.length}건)</td>
                        <td className="text-right">{formatKRW(grandTotal)}</td>
                        <td className="text-right text-neutral-600 text-xs font-medium">{formatKRW(grandSupply)}</td>
                        <td className="text-right text-neutral-600 text-xs font-medium">{formatKRW(grandVat)}</td>
                        <td></td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 서명란 (인쇄용) */}
                <div className="hidden print:flex justify-end mt-12 gap-16">
                  <div className="text-center">
                    <div className="w-24 border-b border-black mb-1 h-10"></div>
                    <p className="text-xs">발행인</p>
                  </div>
                  <div className="text-center">
                    <div className="w-24 border-b border-black mb-1 h-10"></div>
                    <p className="text-xs">수령인</p>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* 인쇄 스타일 */}
      <style jsx global>{`
        @media print {
          /* 기본 레이아웃 숨김 */
          aside, nav, .print\\:hidden { display: none !important; }
          /* A4 설정 */
          @page { size: A4; margin: 15mm 20mm; }
          body { font-size: 11px; color: #000; }
          .print-area { padding: 0; }
          /* 카드 스타일 리셋 */
          .print-area [class*="card"] {
            border: none; box-shadow: none; padding: 0;
          }
        }
      `}</style>
    </>
  );
}
