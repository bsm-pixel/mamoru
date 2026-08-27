'use client';

import { useQuery, useMutation, useQueryClient } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import type { OfflineSale, OfflineSaleItem, Product } from '@/lib/supabase/types';
import toast from 'react-hot-toast';
// 075: cross-domain invalidation 일원화 — mutation 후 대시보드/통계가 즉각 갱신되도록
import { invalidateFinancialQueries } from '@/lib/query/invalidate-keys';

/** 판매 탭 타입 */
export type SalesTab = 'all' | 'today' | 'unpaid' | 'processing' | 'cancelled';
export type SalesChannel = 'all' | 'offline' | 'online' | 'talk' | 'b2b';
export type SalesDateRange = 'all' | 'today' | 'week' | 'month';

/** 오프라인 판매 목록 */
export function useSales(filters?: {
  search?: string;
  page?: number;
  limit?: number;
  tab?: SalesTab;
  channel?: SalesChannel;
  dateRange?: SalesDateRange;
  dateFrom?: string;
  dateTo?: string;
}) {
  const supabase = createClient();
  const page = filters?.page || 1;
  const limit = filters?.limit || 20;
  const from = (page - 1) * limit;
  const to = from + limit - 1;

  return useQuery({
    queryKey: ['sales', filters],
    staleTime: 30_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let query = (supabase as any)
        .from('offline_sales')
        .select('*', { count: 'exact' })
        .order('sale_date', { ascending: false })
        .order('created_at', { ascending: false })
        .range(from, to);

      // 탭 필터 — 전체/오늘/처리필요/미수금은 '종결(취소·반품)' 건 제외, '취소·반품' 탭만 표시
      //   (2026-08-02) 반품(returned_at)도 취소와 동일한 종결 상태로 취급 — 처리필요/미수금에 안 뜨게.
      const tab = filters?.tab || 'all';
      if (tab === 'today') {
        const today = new Date().toISOString().slice(0, 10);
        query = query.eq('sale_date', today).is('cancelled_at', null).is('returned_at', null);
      } else if (tab === 'unpaid') {
        query = query.in('payment_status', ['unpaid', 'partial']).is('cancelled_at', null).is('returned_at', null);
      } else if (tab === 'processing') {
        // 처리 필요 = 결제완료인데 아직 출고·수령 처리 안 함 (사장님 손 필요) — 취소·반품 제외
        query = query.eq('payment_status', 'paid').is('shipped_at', null).is('delivered_at', null).is('cancelled_at', null).is('returned_at', null);
      } else if (tab === 'cancelled') {
        // '취소·반품' 탭 — 취소 또는 반품 건
        query = query.or('cancelled_at.not.is.null,returned_at.not.is.null');
      } else {
        // all: 취소·반품 건 제외 (활성 판매만)
        query = query.is('cancelled_at', null).is('returned_at', null);
      }

      // 채널 필터
      if (filters?.channel && filters.channel !== 'all') {
        if (filters.channel === 'b2b') {
          query = query.in('customer_type', ['dealer', 'academy']);
        } else {
          query = query.eq('sale_channel', filters.channel);
        }
      }

      // 기간 필터: 커스텀 날짜 범위 우선, 없으면 프리셋
      if (filters?.dateFrom) {
        query = query.gte('sale_date', filters.dateFrom);
      } else if (filters?.dateRange && filters.dateRange !== 'all') {
        // ⚠️ 로컬(KST) 달력 기준 — toISOString()(UTC) 쓰면 날짜가 하루/한달 밀림(이번달이 5월까지 나오던 버그)
        const now = new Date();
        const pad = (n: number) => String(n).padStart(2, '0');
        const ymd = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
        let df: string;
        if (filters.dateRange === 'today') {
          df = ymd(now);                              // 오늘 (로컬)
        } else if (filters.dateRange === 'week') {
          const d = new Date(now);
          const dow = d.getDay();                     // 0=일
          d.setDate(d.getDate() + (dow === 0 ? -6 : 1 - dow)); // 이번주 월요일
          df = ymd(d);
        } else {
          df = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-01`; // 이번달 1일
        }
        query = query.gte('sale_date', df);
      }
      if (filters?.dateTo) {
        query = query.lte('sale_date', filters.dateTo);
      }

      // 검색
      if (filters?.search) {
        query = query.or(
          `customer_name.ilike.%${filters.search}%,customer_phone.ilike.%${filters.search}%,sale_number.ilike.%${filters.search}%`
        );
      }

      const { data, count, error } = await query;
      if (error) throw error;
      // 그리드 '채널' 컬럼은 sale_channel 을 직접 표시 → 상담유형 배치조회 불필요(2026-07-17 채널 4분류 개편)
      return { sales: (data || []) as OfflineSale[], total: count || 0 };
    },
  });
}

/** 판매 탭별 건수 */
export function useSalesTabCounts() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sales-tab-counts'],
    staleTime: 30_000,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const today = new Date().toISOString().slice(0, 10);

      const [allRes, todayRes, unpaidRes, processingRes, cancelledRes] = await Promise.all([
        db.from('offline_sales').select('*', { count: 'exact', head: true }).is('cancelled_at', null).is('returned_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).eq('sale_date', today).is('cancelled_at', null).is('returned_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).in('payment_status', ['unpaid', 'partial']).is('cancelled_at', null).is('returned_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).eq('payment_status', 'paid').is('shipped_at', null).is('delivered_at', null).is('cancelled_at', null).is('returned_at', null),
        db.from('offline_sales').select('*', { count: 'exact', head: true }).or('cancelled_at.not.is.null,returned_at.not.is.null'),
      ]);

      return {
        all: allRes.count || 0,
        today: todayRes.count || 0,
        unpaid: unpaidRes.count || 0,
        processing: processingRes.count || 0,
        cancelled: cancelledRes.count || 0,
      };
    },
  });
}

/** 판매 통계 (주간/월간 매출·건수·미수금) */
export function useSalesStats() {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sales-stats'],
    staleTime: 30_000, // 075: 60s → 30s (대시보드 즉각 반영)
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const now = new Date();
      // 로컬(KST) 기준 ISO date — toISOString은 UTC라 KST 자정 직후 4/30로 잘못 변환되는 버그 회피
      const pad2 = (n: number) => String(n).padStart(2, '0');
      const today = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())}`;

      // 이번주 월요일
      const dayOfWeek = now.getDay();
      const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      const monday = new Date(now);
      monday.setDate(now.getDate() + mondayOffset);
      const weekStart = `${monday.getFullYear()}-${pad2(monday.getMonth() + 1)}-${pad2(monday.getDate())}`;

      // 이번달 1일
      const monthStart = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-01`;

      // 2026-05-26 IA 통합: customer_type 도 select — 화면에서 고객(B2C) / 거래처(B2B) 영역별 분리 집계
      // D(2026-08-01 반품시점 반영): 판매는 sale_date(판매월)에 인식(returned 포함), 반품은 아래 *Returns(returned_at) 로 −차감.
      const { data: weekSales } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount, customer_type')
        .gte('sale_date', weekStart)
        .lte('sale_date', today)
        .is('cancelled_at', null);

      // 월간 매출
      const { data: monthSales } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount, customer_type')
        .gte('sale_date', monthStart)
        .lte('sale_date', today)
        .is('cancelled_at', null);

      // 이번달 복원수리(RS) 항목 — 카드 '제품/복원수리' 분리표기용. offline_sales 와 동일 필터(월·취소)로 조인.
      const { data: monthRsItems } = await db
        .from('offline_sale_items')
        .select('total_price, offline_sales!inner(customer_type)')
        .eq('category', 'RS')
        .gt('total_price', 0)
        .gte('offline_sales.sale_date', monthStart)
        .lte('offline_sales.sale_date', today)
        .is('offline_sales.cancelled_at', null);

      // ─── D 반품(−): returned_at 이 이 주/이번달인 판매·RS (반품시점 반영) ───
      const wkFrom = `${weekStart}T00:00:00`, moFrom = `${monthStart}T00:00:00`, toEnd = `${today}T23:59:59`;
      const { data: weekReturns } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount, customer_type')
        .gte('returned_at', wkFrom).lte('returned_at', toEnd).is('cancelled_at', null);
      const { data: monthReturns } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount, customer_type')
        .gte('returned_at', moFrom).lte('returned_at', toEnd).is('cancelled_at', null);
      const { data: monthRsReturns } = await db
        .from('offline_sale_items')
        .select('total_price, offline_sales!inner(customer_type)')
        .eq('category', 'RS').gt('total_price', 0)
        .gte('offline_sales.returned_at', moFrom).lte('offline_sales.returned_at', toEnd)
        .is('offline_sales.cancelled_at', null);

      // 미수금 총액 (취소·반품 제외 — 반품건은 더 이상 받을 돈 아님)
      const { data: unpaidSales } = await db
        .from('offline_sales')
        .select('total_amount, paid_amount, discount_amount, customer_type')
        .in('payment_status', ['unpaid', 'partial'])
        .is('cancelled_at', null)
        .is('returned_at', null);

      // 발생주의: 매출 = total_amount - discount_amount (미수금 포함)
      type SalesRow = { total_amount: number; paid_amount: number; discount_amount: number; customer_type?: string };
      const isPartner = (t?: string) => t === 'dealer' || t === 'academy';
      const sum = (arr: SalesRow[] | null) => {
        if (!arr) return { amount: 0, count: 0 };
        const amount = arr.reduce((s, r) => s + ((r.total_amount || 0) - (r.discount_amount || 0)), 0);
        return { amount, count: arr.length };
      };
      const sumOutstanding = (arr: SalesRow[] | null) => {
        if (!arr) return 0;
        return arr.reduce(
          (s, r) => s + ((r.total_amount - (r.discount_amount || 0)) - (r.paid_amount || 0)),
          0
        );
      };
      const partition = (arr: SalesRow[] | null): { customer: SalesRow[]; partner: SalesRow[] } => {
        const customer: SalesRow[] = [];
        const partner: SalesRow[] = [];
        for (const r of arr || []) {
          if (isPartner(r.customer_type)) partner.push(r);
          else customer.push(r);
        }
        return { customer, partner };
      };

      const weekParts = partition(weekSales);
      const monthParts = partition(monthSales);
      const unpaidParts = partition(unpaidSales);
      // D: 반품(returned_at) partition — amount 만 −차감 (count 는 판매월 기준 유지)
      const weekRetParts = partition(weekReturns);
      const monthRetParts = partition(monthReturns);
      const net = (salesArr: SalesRow[] | null, retArr: SalesRow[] | null) =>
        ({ amount: sum(salesArr).amount - sum(retArr).amount, count: (salesArr || []).length });

      const unpaidTotal = sumOutstanding(unpaidSales);

      // 이번달 RS 합을 B2C/B2B 로 분리 (판매월 + / 반품월 −)
      type RsRow = { total_price: number; offline_sales?: { customer_type?: string } | null };
      let customerMonthRepair = 0;
      let partnerMonthRepair = 0;
      for (const r of (monthRsItems || []) as RsRow[]) {
        const amt = r.total_price || 0;
        if (isPartner(r.offline_sales?.customer_type)) partnerMonthRepair += amt;
        else customerMonthRepair += amt;
      }
      for (const r of (monthRsReturns || []) as RsRow[]) {  // D: 반품 RS −차감
        const amt = r.total_price || 0;
        if (isPartner(r.offline_sales?.customer_type)) partnerMonthRepair -= amt;
        else customerMonthRepair -= amt;
      }

      // B2B 거래처별 이번달 매출 (판매월 +) — 반품은 아래 b2bReturns 로 −차감
      const { data: b2bSales } = await db
        .from('offline_sales')
        .select('customer_name, customer_type, total_amount, discount_amount')
        .in('customer_type', ['dealer', 'academy'])
        .gte('sale_date', monthStart)
        .lte('sale_date', today)
        .is('cancelled_at', null);
      const { data: b2bReturns } = await db
        .from('offline_sales')
        .select('customer_name, customer_type, total_amount, discount_amount')
        .in('customer_type', ['dealer', 'academy'])
        .gte('returned_at', moFrom).lte('returned_at', toEnd)
        .is('cancelled_at', null);

      const b2bMap: Record<string, { name: string; type: string; amount: number; count: number }> = {};
      for (const s of (b2bSales || [])) {
        const key = s.customer_name || '미지정';
        if (!b2bMap[key]) b2bMap[key] = { name: key, type: s.customer_type || '', amount: 0, count: 0 };
        b2bMap[key].amount += (s.total_amount || 0) - (s.discount_amount || 0);
        b2bMap[key].count++;
      }
      for (const s of (b2bReturns || [])) {  // D: 반품 −차감 (건수는 판매월 기준 유지)
        const key = s.customer_name || '미지정';
        if (!b2bMap[key]) b2bMap[key] = { name: key, type: s.customer_type || '', amount: 0, count: 0 };
        b2bMap[key].amount -= (s.total_amount || 0) - (s.discount_amount || 0);
      }
      const b2bRanking = Object.values(b2bMap).sort((a, b) => b.amount - a.amount);

      return {
        // 기존 필드 — 전체 합계 (B2C+B2B 합산). D: 반품시점 −차감(net)
        week: net(weekSales, weekReturns),
        month: net(monthSales, monthReturns),
        outstanding: unpaidTotal,
        b2b: b2bRanking,
        // 2026-05-26 신규 — 영역별 분리 (offline_sales 기준만, deliveries 합산은 화면에서). D: net
        customerWeek: net(weekParts.customer, weekRetParts.customer),
        customerMonth: net(monthParts.customer, monthRetParts.customer),
        customerOutstanding: sumOutstanding(unpaidParts.customer),
        partnerWeek: net(weekParts.partner, weekRetParts.partner),
        partnerMonth: net(monthParts.partner, monthRetParts.partner),
        partnerOutstanding: sumOutstanding(unpaidParts.partner),
        // 2026-06-18 제품/복원수리 분리표기 — 이번달 RS 합(offline_sales 기준). 제품 = customerMonth.amount − customerMonthRepair
        customerMonthRepair,
        partnerMonthRepair,
      };
    },
  });
}

/** 오프라인 판매 단건 조회 (항목 + 시리얼 + 070: 원본 상담 link 포함) */
export function useSale(id: string) {
  const supabase = createClient();

  return useQuery({
    queryKey: ['sale', id],
    queryFn: async () => {
      const [saleRes, itemsRes, serialsRes] = await Promise.all([
        supabase.from('offline_sales').select('*').eq('id', id).single(),
        supabase.from('offline_sale_items').select('*').eq('sale_id', id),
        supabase.from('product_serials').select('id, serial_number, product_id, sale_item_id').eq('offline_sale_id', id),
      ]);
      if (saleRes.error) throw saleRes.error;
      const sale = saleRes.data as OfflineSale;

      // customer_id가 있으면 최신 연락처 + 매장명/활동명/직급 가져오기
      let customerInfo: { company_name: string | null; activity_name: string | null; position: string | null; postcode: string | null; address_road: string | null; address_detail: string | null } | null = null;
      if (sale.customer_id) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data: cust } = await (supabase as any)
          .from('customers')
          .select('phone, company_name, activity_name, position, postcode, address_road, address_detail')
          .eq('id', sale.customer_id)
          .single();
        if (cust?.phone) {
          (sale as Record<string, unknown>).customer_phone = cust.phone;
        }
        if (cust) {
          customerInfo = {
            company_name: cust.company_name ?? null,
            activity_name: cust.activity_name ?? null,
            position: cust.position ?? null,
            postcode: cust.postcode ?? null,          // 주소지 표시용 (2026-07-14)
            address_road: cust.address_road ?? null,
            address_detail: cust.address_detail ?? null,
          };
        }
      }

      // 070: source_consultation_id 있으면 원본 상담 정보 조회 (mirror 모드용)
      let linkedConsultation: { id: string; unique_id: string; name: string; status: string; consultation_type: string } | null = null;
      const sourceId = (sale as { source_consultation_id?: string | null }).source_consultation_id;
      if (sourceId) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const { data: consult } = await (supabase as any)
          .from('consultations')
          .select('id, unique_id, name, status, consultation_type')
          .eq('id', sourceId)
          .single();
        if (consult) linkedConsultation = consult;
      }

      return {
        sale,
        items: (itemsRes.data || []) as OfflineSaleItem[],
        // sale_item_id 포함 — 상세 페이지에서 정확한 시리얼-상품 매칭에 사용 (2026-05-17 fix)
        serials: (serialsRes.data || []) as Array<{ id: string; serial_number: string; product_id: string | null; sale_item_id: string | null }>,
        linkedConsultation,
        customerInfo,
      };
    },
    enabled: !!id,
  });
}

/** 제품 목록 (판매 입력용) */
export function useProducts(opts?: { includeInactive?: boolean }) {
  const supabase = createClient();
  const includeInactive = opts?.includeInactive || false;

  return useQuery({
    queryKey: ['products', { includeInactive }],
    staleTime: 60 * 1000, // 075: 5분 → 1분 (시리얼 sold/재고 변경 후 빠른 반영)
    queryFn: async () => {
      let query = supabase
        .from('products')
        .select('*')
        .order('category')
        .order('name');
      if (!includeInactive) query = query.eq('is_active', true);
      const { data, error } = await query;
      if (error) throw error;
      return (data || []) as Product[];
    },
  });
}

/** 오프라인 판매 생성 */
export function useCreateSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async (data: {
      sale: {
        customer_id?: string;
        customer_name: string;
        customer_phone?: string;
        sale_date?: string;
        total_amount: number;
        discount_amount?: number;
        paid_amount: number;
        payment_method: string;
        payment_status?: string;
        payment_detail?: Record<string, number>;
        customer_type?: string;
        memo?: string;
        sale_channel?: string;
        contract_id?: string | null;
        source_consultation_id?: string;        // 070: 상담 → 판매 link
      };
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        serial_ids?: string[];
      }>;
      /** Phase A — 시리얼 다른 판매에서 이전 명시 동의 (2026-05-18) */
      allow_serial_transfer?: boolean;
    }) => {
      const res = await fetch('/api/sales', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '판매 등록 실패');
      }
      return res.json();
    },
    onSuccess: (data) => {
      toast.success(`판매 등록 완료: ${data.saleNumber}`);
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('판매 등록 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 판매 취소 — 서버 확인 필수 (시리얼/재고/아임웹 역전) */
export function useCancelSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, reason }: { id: string; reason?: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'cancel', reason }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '판매 취소 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('판매가 취소되었습니다');
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('판매 취소 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 반품 처리 */
export function useReturnSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, reason }: { id: string; reason?: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'return', reason }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '반품 처리 실패');
      }
      return res.json();
    },
    onSuccess: () => {
      toast.success('반품 처리 완료 — 재고/시리얼 복원됨');
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('반품 처리 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 결제상태 변경 — 낙관적 업데이트 */
export function useUpdatePaymentStatus() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, payment_status, paid_amount }: {
      id: string;
      payment_status: string;
      paid_amount?: number;
    }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'update_payment', payment_status, paid_amount }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '결제상태 변경 실패');
      }
      return res.json();
    },
    onMutate: async ({ id, payment_status, paid_amount }) => {
      await queryClient.cancelQueries({ queryKey: ['sale', id] });

      const prevDetail = queryClient.getQueryData(['sale', id]);

      // 상세 캐시 즉시 업데이트
      queryClient.setQueryData(['sale', id], (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const data = old as { sale: OfflineSale; items: OfflineSaleItem[]; serials: unknown[] };
        return {
          ...data,
          sale: {
            ...data.sale,
            payment_status,
            ...(paid_amount !== undefined ? { paid_amount } : {}),
          },
        };
      });

      return { prevDetail };
    },
    onError: (err, { id }, context) => {
      if (context?.prevDetail) queryClient.setQueryData(['sale', id], context.prevDetail);
      toast.error('결제상태 변경 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
    onSuccess: () => {
      toast.success('결제상태가 변경되었습니다');
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
  });
}

/** 메모 수정 — 낙관적 업데이트 */
export function useUpdateSaleMemo() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, memo }: { id: string; memo: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'update_memo', memo }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || '메모 수정 실패');
      }
      return res.json();
    },
    onMutate: async ({ id, memo }) => {
      await queryClient.cancelQueries({ queryKey: ['sale', id] });
      const prevDetail = queryClient.getQueryData(['sale', id]);

      queryClient.setQueryData(['sale', id], (old: unknown) => {
        if (!old || typeof old !== 'object') return old;
        const data = old as { sale: OfflineSale; items: OfflineSaleItem[]; serials: unknown[] };
        return { ...data, sale: { ...data.sale, memo } };
      });

      return { prevDetail };
    },
    onError: (err, { id }, context) => {
      if (context?.prevDetail) queryClient.setQueryData(['sale', id], context.prevDetail);
      toast.error('메모 수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
    onSettled: (_d, _e, { id }) => {
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
    },
  });
}

/** 판매 정보 수정 (금액/할인/결제방법/날짜) */
export function useEditSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, ...fields }: {
      id: string;
      total_amount?: number;
      discount_amount?: number;
      payment_method?: string;
      sale_date?: string;
      payment_detail?: Record<string, number>;
    }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'edit_sale', ...fields }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '수정 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('판매 정보가 수정되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 판매 재구성 (제품 추가/삭제 — 시리얼/재고 자동 재조정) */
export function useRebuildSale() {
  const queryClient = useQueryClient();

  return useMutation({
    mutationFn: async ({ id, items, sale_info, allow_serial_transfer, exchange_returned_serial_ids, exchange_return_nonserial }: {
      id: string;
      items: Array<{
        product_id?: string;
        product_name: string;
        sku?: string;
        quantity: number;
        unit_price: number;
        total_price: number;
        serial_ids?: string[];
        manual_serials?: string[];
      }>;
      sale_info: {
        total_amount: number;
        discount_amount?: number;
        payment_method: string;
        payment_status?: string;
        paid_amount?: number;
        sale_date?: string;
        payment_detail?: Record<string, number>;
        memo?: string;
        sale_channel?: string;
      };
      /** Phase A — 시리얼 다른 판매에서 이전 명시 동의 (2026-05-18) */
      allow_serial_transfer?: boolean;
      /** 🔁 교환 — 제거 품목의 시리얼 id(반품창고行). 비우면 일반 수정 */
      exchange_returned_serial_ids?: string[];
      /** 🔁 교환(비시리얼) — { product_id: qty } 제거 비시리얼 수량을 return_stock으로 */
      exchange_return_nonserial?: Record<string, number>;
    }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'rebuild_sale', items, sale_info, allow_serial_transfer, exchange_returned_serial_ids, exchange_return_nonserial }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '수정 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('판매가 수정되었습니다 (시리얼/재고 재조정 완료)');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => {
      toast.error('판매 수정 실패: ' + (err instanceof Error ? err.message : String(err)));
    },
  });
}

/** 판매 송장 생성 (ALPS) */
export function useShipSale() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/sales/${id}/ship`, { method: 'POST' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '송장 생성 실패');
      }
      return res.json();
    },
    onSuccess: (data, id) => {
      toast.success(`송장 생성 완료: ${data.invoiceNumber}`);
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 판매 송장 취소 */
export function useCancelSaleShipment() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async (id: string) => {
      const res = await fetch(`/api/sales/${id}/ship`, { method: 'DELETE' });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '송장 취소 실패');
      }
      return res.json();
    },
    onSuccess: (data, id) => {
      // ALPS 집하취소는 API 미지원 → 실패 시 warning 을 반드시 노출(거짓 성공 신호 방지). 복원수리·빠른송장과 동일 패턴
      if (data?.warning) toast(data.warning, { icon: '⚠️', duration: 8000 });
      else toast.success('송장 취소 완료');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 판매 출고완료 처리 (shipped_at 설정 + 선택적 알림톡) */
export function useMarkSaleShipped() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, send_notification }: { id: string; send_notification: boolean }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'mark_shipped', send_notification }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '출고완료 처리 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('출고완료 처리되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 포장완료(준비완료) 토글 — 2026-07-18
 *  packed_at 만 기록/해제하는 내부 표시용. 알림톡·외부연동 없음.
 *  ids 배열을 받아 단건([id])·일괄(여러 건) 모두 같은 훅으로 처리. */
export function useMarkSalePacked() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ ids, packed }: { ids: string[]; packed: boolean }) => {
      const action = packed ? 'mark_packed' : 'unmark_packed';
      const results = await Promise.all(
        ids.map(async (id) => {
          const res = await fetch(`/api/sales/${id}`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ action }),
          });
          if (!res.ok) {
            const err = await res.json().catch(() => ({ error: res.statusText }));
            return { id, ok: false, error: typeof err.error === 'string' ? err.error : '처리 실패' };
          }
          return { id, ok: true };
        }),
      );
      return results;
    },
    onSuccess: (results, { ids, packed }) => {
      const okCount = results.filter((r) => r.ok).length;
      const failed = results.filter((r) => !r.ok);
      if (failed.length === 0) {
        toast.success(packed ? `준비완료 처리 (${okCount}건)` : '준비완료를 해제했습니다');
      } else {
        // 일부 실패해도 성공분은 반영됨 — 조용히 넘기지 않고 이유를 보여준다
        toast.error(`${okCount}/${ids.length}건 처리 · 실패: ${failed[0].error}`);
      }
      ids.forEach((id) => queryClient.invalidateQueries({ queryKey: ['sale', id] }));
      queryClient.invalidateQueries({ queryKey: ['sales'] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/** 출고 알림톡 수동 재발송 (2026-07-15) — 이미 출고됐는데 알림톡이 안 나간 B2C 건 */
export function useResendShipNotify() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id }: { id: string }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'resend_ship_notify' }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : '알림톡 발송 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id }) => {
      toast.success('출고 알림톡을 발송했습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

/**
 * 판매 배송완료 처리 — 두 경로 통합 (2026-05-25)
 *   1. mode='delivery': 송장 있는 택배 발송 케이스 — ALPS 추적 실패 fallback (수동 배송완료)
 *   2. mode='pickup':   송장 없는 매장 직접 수령 케이스 — "고객 수령 완료" 처리
 *   두 경로 모두 offline_sales.delivered_at 설정. invoice_number 유무로 칩 표시 분기.
 */
export function useMarkSaleDelivered() {
  const queryClient = useQueryClient();
  return useMutation({
    mutationFn: async ({ id, mode }: { id: string; mode: 'delivery' | 'pickup' }) => {
      const res = await fetch(`/api/sales/${id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'mark_delivered', mode }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(typeof err.error === 'string' ? err.error : JSON.stringify(err.error) || '배송완료 처리 실패');
      }
      return res.json();
    },
    onSuccess: (_d, { id, mode }) => {
      toast.success(mode === 'pickup' ? '고객 수령 완료 처리되었습니다' : '배송완료 처리되었습니다');
      queryClient.invalidateQueries({ queryKey: ['sale', id] });
      queryClient.invalidateQueries({ queryKey: ['sales'] });
      invalidateFinancialQueries(queryClient);
    },
    onError: (err) => toast.error(err instanceof Error ? err.message : String(err)),
  });
}

