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

export interface MarginSummary {
  total_cogs: number;
  gross_profit: number;
  margin_rate: number;
}

export interface ProductMargin {
  product_id: string;
  product_name: string;
  sku: string;
  category?: string | null;
  qty: number;
  revenue: number;
  cogs: number;
  profit: number;
  margin_rate: number;
  b2c_qty?: number;
  b2c_revenue?: number;
  b2c_cogs?: number;
  b2b_qty?: number;
  b2b_revenue?: number;
  b2b_cogs?: number;
}

export interface SupplierSpend {
  name: string;
  total: number;
  count: number;
}

export interface PayableItem {
  name: string;
  total_owed: number;
  count: number;
}

export interface ReceivableItem {
  id: string;
  name: string;
  outstanding: number;
}

/** 복원수리 매출 = A(접수 repairs) + B(판매시스템 RS) + C(납품 RS) */
export interface RepairSalesSummary {
  total: number;            // A + B + C
  a_intake: number;         // 접수시스템 (repairs, 입금 기준 실제 금액)
  b_offline_rs: number;     // 판매시스템 RS (offline_sale_items category='RS')
  c_delivery_rs: number;    // 납품 RS (delivery_items category='RS')
  a_count: number;
  b_count: number;
  c_count: number;
  bc_qty: number;           // B+C 자루 수
  service_cost_total: number; // (A 기준)
  shipping_fee_total: number; // (A 기준)
  count: number;            // 전체 건수 (호환)
}

/** 제품 매출 = B2C(소매/온라인) + B2B(딜러/아카데미 + 납품) — RS 제외 */
export interface ProductSalesSummary {
  total: number;
  b2c: number;
  b2b: number;
  b2b_offline: number;
  b2b_delivery: number;
  offline_count: number;
  delivery_count: number;
}

export interface ReportData {
  period: { from: string; to: string };
  sales: SaleSummary;             // 오프라인 판매 전체 (RS 포함 — 호환 "오프라인 판매 총액")
  product_sales?: ProductSalesSummary;  // 제품 매출 (RS 제외, 납품 포함)
  repair_sales?: RepairSalesSummary;
  total_revenue?: number;
  purchases: PurchaseSummary;
  vat: VATSummary;
  margin: MarginSummary;
  by_product: ProductMargin[];
  by_supplier: SupplierSpend[];
  payables: { items: PayableItem[]; total: number };
  receivables: { items: ReceivableItem[]; total: number };
  daily: {
    sales: Record<string, number>;
    purchases: Record<string, number>;
    repairs?: Record<string, number>;
    product?: Record<string, number>; // 제품 일별(판매−RS + 납품제품) — 탭별 차트용
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
