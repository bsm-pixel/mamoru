'use client';

import { useQuery } from '@tanstack/react-query';

export interface SaleSummary {
  count: number;
  total: number;
  supply: number;
  vat: number;
  discount: number;
  paid: number;
  by_method: Record<string, number>;
}

export interface PurchaseSummary {
  count: number;
  total: number;
  supply: number;
  vat: number;
  deposit: number;
}

export interface VATSummary {
  sales_vat: number;
  purchase_vat: number;
  net_vat: number;
}

export interface SaleDetail {
  id: string;
  sale_date: string;
  customer_name: string;
  customer_id: string | null;
  total_amount: number;
  supply_amount: number;
  vat_amount: number;
  payment_method: string;
  payment_status: string;
  discount_amount: number;
  paid_amount: number;
}

export interface PurchaseDetail {
  id: string;
  order_date: string;
  supplier_name: string;
  total_amount: number;
  supply_amount: number;
  vat_amount: number;
  deposit_amount: number;
  balance_amount: number;
  status: string;
}

export interface ReportData {
  period: { from: string; to: string };
  sales: SaleSummary;
  purchases: PurchaseSummary;
  vat: VATSummary;
  daily: {
    sales: Record<string, number>;
    purchases: Record<string, number>;
  };
  details: {
    sales: SaleDetail[];
    purchases: PurchaseDetail[];
  };
}

/** 매출/매입/VAT 기간별 요약 */
export function useReportSummary(from: string, to: string) {
  return useQuery<ReportData>({
    queryKey: ['reports', 'summary', from, to],
    queryFn: async () => {
      const res = await fetch(`/api/reports/summary?from=${from}&to=${to}`);
      if (!res.ok) throw new Error(await res.text());
      return res.json();
    },
    enabled: !!from && !!to,
  });
}

/** 엑셀 다운로드 트리거 */
export function downloadExcel(type: 'sales' | 'purchases', from: string, to: string) {
  const url = `/api/reports/export?type=${type}&from=${from}&to=${to}`;
  window.open(url, '_blank');
}
