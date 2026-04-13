'use client';

import { useState, useEffect, memo } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Pagination } from '@/components/ui/pagination';
import { SlidePanel } from '@/components/ui/slide-panel';
import { ConfirmModal } from '@/components/ui/confirm-modal';
import { DeliveryDetailPanel } from '@/components/deliveries/delivery-detail-panel';
import { useDeliveries, useDeliveryStats, useCreateDelivery } from '@/hooks/use-deliveries';
import { useProducts } from '@/hooks/use-sales';
import { useCustomerSearch } from '@/hooks/use-customers';
import { useSetting } from '@/hooks/use-settings';
import { formatKRW, formatDate, formatPhone, calcVAT } from '@/lib/utils/format';
import { Package, Plus, X, AlertCircle, Calendar, TrendingUp } from 'lucide-react';
import toast from 'react-hot-toast';
import type { Product } from '@/lib/supabase/types';

/* ── 상수 ── */
const STATUS_LABEL: Record<string, string> = {
  draft: '작성중', confirmed: '납품확정', shipped: '출고완료', settled: '정산완료',
};
const STATUS_COLOR: Record<string, string> = {
  draft: 'bg-neutral-100 text-neutral-600',
  confirmed: 'bg-blue-100 text-blue-700',
  shipped: 'bg-green-100 text-green-700',
  settled: 'bg-emerald-100 text-emerald-700',
};
const PAYMENT_LABEL: Record<string, string> = { unpaid: '미결제', partial: '부분결제', paid: '결제완료' };
const PAYMENT_COLOR: Record<string, string> = { unpaid: 'bg-red-100 text-red-600', partial: 'bg-yellow-100 text-yellow-700', paid: 'bg-green-100 text-green-700' };
const RECEIPT_LABEL: Record<string, string> = { expense_proof: '지출증빙', tax_invoice: '세금계산서', none: '미적용' };

const STATUS_TABS = [
  { value: '', label: '전체' },
  { value: 'draft', label: '작성중' },
  { value: 'confirmed', label: '납품확정' },
  { value: 'shipped', label: '출고완료' },
  { value: 'settled', label: '정산완료' },
];

const DATE_RANGES = [
  { value: 'all', label: '전체 기간' },
  { value: 'today', label: '오늘' },
  { value: 'week', label: '이번주' },
  { value: 'month', label: '이번달' },
] as const;

export default function DeliveriesPage() {
  const [statusFilter, setStatusFilter] = useState('');
  const [search, setSearch] = useState('');
  const [dateRange, setDateRange] = useState<string>('all');
  const [page, setPage] = useState(1);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [showCreate, setShowCreate] = useState(false);
  const limit = 20;

  const [isLg, setIsLg] = useState(false);
  useEffect(() => {
    const mql = window.matchMedia('(min-width: 1024px)');
    setIsLg(mql.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mql.addEventListener('change', handler);
    return () => mql.removeEventListener('change', handler);
  }, []);

  const { data, isLoading } = useDeliveries({
    status: statusFilter || undefined,
    search: search || undefined,
    dateRange,
    page,
    limit,
  });
  const deliveries = data?.deliveries || [];
  const total = data?.total || 0;
  const totalPages = Math.ceil(total / limit);
  const { data: stats } = useDeliveryStats();

  /* ── 목록 영역 ── */
  const listContent = (
    <>
      {/* 통계 카드 */}
      {stats && (
        <div className="grid grid-cols-3 gap-2">
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <Calendar size={12} />이번주
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.weekAmount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.weekCount}건</p>
          </div>
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-neutral-500 mb-1">
              <TrendingUp size={12} />이번달
            </div>
            <p className="text-base font-bold text-neutral-900">{formatKRW(stats.monthAmount)}</p>
            <p className="text-[11px] text-neutral-400">{stats.monthCount}건</p>
          </div>
          <div className="bg-white rounded-lg border border-neutral-200 p-3">
            <div className="flex items-center gap-1.5 text-xs text-red-500 mb-1">
              <AlertCircle size={12} />미수금
            </div>
            <p className="text-base font-bold text-red-600">{formatKRW(stats.outstanding)}</p>
          </div>
        </div>
      )}

      {/* 검색 + 생성 버튼 */}
      <div className="flex items-center gap-2">
        <Button size="sm" onClick={() => setShowCreate(true)}>
          <Plus size={14} />납품서 작성
        </Button>
        <SearchInput
          value={search}
          onChange={(v) => { setSearch(v); setPage(1); }}
          placeholder="납품번호, 고객명 검색"
        />
      </div>

      {/* 탭 바 */}
      <div className="flex gap-1 border-b border-neutral-200">
        {STATUS_TABS.map((tab) => {
          const count = stats ? (tab.value === '' ? stats.all : stats[tab.value as keyof typeof stats] as number) : 0;
          const active = statusFilter === tab.value;
          return (
            <button
              key={tab.value}
              onClick={() => { setStatusFilter(tab.value); setPage(1); }}
              className={`px-3 py-2 text-sm font-medium border-b-2 transition whitespace-nowrap ${
                active
                  ? 'border-neutral-900 text-neutral-900'
                  : 'border-transparent text-neutral-400 hover:text-neutral-600'
              }`}
            >
              {tab.label}
              {typeof count === 'number' && count > 0 && (
                <span className={`ml-1.5 text-xs px-1.5 py-0.5 rounded-full ${
                  active ? 'bg-neutral-900 text-white' : 'bg-neutral-200 text-neutral-500'
                }`}>
                  {count}
                </span>
              )}
            </button>
          );
        })}
      </div>

      {/* 기간 필터 */}
      <div className="flex gap-1">
        {DATE_RANGES.map((d) => (
          <button
            key={d.value}
            onClick={() => { setDateRange(d.value); setPage(1); }}
            className={`px-2.5 py-1 text-xs rounded-full border transition ${
              dateRange === d.value
                ? 'bg-neutral-900 text-white border-neutral-900'
                : 'bg-white text-neutral-500 border-neutral-200 hover:border-neutral-400'
            }`}
          >
            {d.label}
          </button>
        ))}
      </div>

      {/* 납품 목록 */}
      <Card padding={false}>
        {isLoading ? (
          <div className="p-4 space-y-3">
            {Array.from({ length: 5 }).map((_, i) => (
              <Skeleton key={i} className="h-14 w-full" />
            ))}
          </div>
        ) : deliveries.length === 0 ? (
          <EmptyState icon={Package} message="납품 내역이 없습니다" />
        ) : (
          <div className="divide-y divide-neutral-100">
            {deliveries.map((dl) => (
              <DeliveryRow
                key={dl.id as string}
                dl={dl}
                isSelected={selectedId === dl.id}
                onClick={() => setSelectedId(dl.id as string)}
              />
            ))}
          </div>
        )}
      </Card>

      <Pagination page={page} totalPages={totalPages} total={total} onPageChange={setPage} />
    </>
  );

  return (
    <>
      <Topbar title="납품관리" />

      {isLg ? (
        /* PC: 마스터-디테일 2컬럼 */
        <div className="flex gap-4 px-4 md:px-6 py-4 h-full min-h-0">
          <div className="w-2/5 shrink-0 overflow-y-auto space-y-3 pr-1">
            {listContent}
          </div>
          <div className="flex-1 min-w-0 overflow-y-auto bg-white rounded-xl border border-neutral-200">
            {selectedId ? (
              <DeliveryDetailPanel deliveryId={selectedId} />
            ) : (
              <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
                <Package size={28} className="mb-2 opacity-40" />
                <p className="text-xs">목록에서 납품 건을 선택하세요</p>
              </div>
            )}
          </div>
        </div>
      ) : (
        /* 모바일: 목록 */
        <div className="px-4 md:px-6 py-4 space-y-4">
          {listContent}
        </div>
      )}

      {/* 모바일 상세 패널 */}
      {!isLg && (
        <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="납품 상세" className="sm:w-[480px]">
          {selectedId && <DeliveryDetailPanel deliveryId={selectedId} />}
        </SlidePanel>
      )}

      {/* 납품서 작성 모달 */}
      {showCreate && (
        <CreateDeliveryModal onClose={() => setShowCreate(false)} onCreated={(id) => { setShowCreate(false); setSelectedId(id); }} />
      )}
    </>
  );
}

/* ── 목록 행 ── */
const DeliveryRow = memo(function DeliveryRow({ dl, isSelected, onClick }: {
  dl: Record<string, unknown>; isSelected: boolean; onClick: () => void;
}) {
  const status = (dl.status as string) || 'draft';
  const paymentStatus = (dl.payment_status as string) || 'unpaid';

  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-4 px-4 py-3 cursor-pointer transition ${
        isSelected ? 'bg-neutral-50 border-l-2 border-l-neutral-900' : 'hover:bg-warm-ivory/60'
      }`}
    >
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-sm font-semibold text-indigo-black truncate">
            {dl.customer_name as string}
          </span>
          <Badge className={STATUS_COLOR[status] || STATUS_COLOR.draft}>
            {STATUS_LABEL[status] || status}
          </Badge>
          <Badge className={PAYMENT_COLOR[paymentStatus] || PAYMENT_COLOR.unpaid}>
            {PAYMENT_LABEL[paymentStatus] || paymentStatus}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
          <span>{dl.dl_number as string}</span>
          <span>{formatDate(dl.delivery_date as string)}</span>
        </div>
      </div>
      <div className="text-right shrink-0">
        {dl.payment_status === 'partial' && (dl.paid_amount as number) > 0 && (
          <p className="text-[10px] text-yellow-600 font-medium">{formatKRW((dl.paid_amount as number))} 선납</p>
        )}
        <p className="text-sm font-bold">{formatKRW((dl.total_amount as number) || 0)}</p>
      </div>
    </div>
  );
});

/* ── 납품서 작성 모달 ── */
function CreateDeliveryModal({ onClose, onCreated }: { onClose: () => void; onCreated: (id: string) => void }) {
  const createDelivery = useCreateDelivery();
  const { data: products = [] } = useProducts();
  const taxType = useSetting<string>('accounting.tax_type', 'simplified');

  // 고객 검색
  const [customerQuery, setCustomerQuery] = useState('');
  const { data: searchResults = [] } = useCustomerSearch(customerQuery);
  const [selectedCustomer, setSelectedCustomer] = useState<{
    id: string; name: string; phone: string | null; customer_type?: string;
  } | null>(null);
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [showCustomerDropdown, setShowCustomerDropdown] = useState(false);

  // 제품
  const [productSearch, setProductSearch] = useState('');
  const [cart, setCart] = useState<Array<{
    product_id?: string; product_name: string; sku?: string; category?: string;
    quantity: number; unit_price: number;
  }>>([]);

  // 결제/옵션
  const [deliveryDate, setDeliveryDate] = useState(new Date().toISOString().slice(0, 10));
  const [expectedDate, setExpectedDate] = useState('');
  const [vatType, setVatType] = useState<'included' | 'separate' | 'none'>('included');
  const [receiptType, setReceiptType] = useState<string>(
    taxType === 'general' ? 'tax_invoice' : 'expense_proof'
  );
  const [paymentStatus, setPaymentStatus] = useState<'unpaid' | 'partial' | 'paid'>('unpaid');
  const [paymentMethod, setPaymentMethod] = useState('transfer');
  const [discount, setDiscount] = useState(0);
  const [paidAmount, setPaidAmount] = useState(0);
  const [memo, setMemo] = useState('');
  const [customName, setCustomName] = useState('');
  const [customPrice, setCustomPrice] = useState('');

  const customerType = selectedCustomer?.customer_type;

  // 제품 필터
  const filteredProducts = products.filter((p) => {
    if (!productSearch) return true;
    const q = productSearch.toLowerCase();
    return p.name.toLowerCase().includes(q) || (p.sku || '').toLowerCase().includes(q);
  });

  // 가격 결정: 고객 유형별
  function getPrice(p: Product): number {
    if (customerType === 'dealer' && (p as Record<string, unknown>).price_dealer) return (p as Record<string, unknown>).price_dealer as number;
    if (customerType === 'academy' && (p as Record<string, unknown>).price_academy) return (p as Record<string, unknown>).price_academy as number;
    return p.price;
  }

  function addProduct(p: Product) {
    const existing = cart.find((c) => c.product_id === p.id);
    if (existing) {
      setCart((prev) => prev.map((c) => c.product_id === p.id ? { ...c, quantity: c.quantity + 1 } : c));
    } else {
      setCart((prev) => [...prev, {
        product_id: p.id,
        product_name: p.name,
        sku: p.sku || undefined,
        category: p.category || undefined,
        quantity: 1,
        unit_price: getPrice(p),
      }]);
    }
  }

  // 합계
  const itemTotal = cart.reduce((s, c) => s + c.quantity * c.unit_price, 0);
  const baseAmount = itemTotal - discount;
  const { supply, vat, payment: totalAmount } = calcVAT(baseAmount, vatType);

  async function handleSubmit() {
    const name = selectedCustomer?.name || customerName.trim();
    if (!name) { toast.error('거래처를 입력해주세요'); return; }
    if (cart.length === 0) { toast.error('품목을 추가해주세요'); return; }

    try {
      const result = await createDelivery.mutateAsync({
        customer_id: selectedCustomer?.id,
        customer_name: name,
        customer_phone: selectedCustomer?.phone || customerPhone.trim() || undefined,
        customer_type: customerType,
        delivery_date: deliveryDate,
        expected_date: expectedDate || undefined,
        memo: memo.trim() || undefined,
        vat_type: vatType,
        receipt_type: receiptType,
        payment_status: paymentStatus,
        payment_method: paymentMethod,
        paid_amount: paymentStatus === 'partial' ? paidAmount : paymentStatus === 'paid' ? totalAmount : 0,
        discount_amount: discount,
        items: cart,
      });
      onCreated((result.delivery as Record<string, unknown>).id as string);
    } catch {
      // error handled by hook
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '780px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="flex items-center justify-between px-5 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">납품서 작성</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600"><X size={18} /></button>
        </div>

        {/* 본문 */}
        <div className="overflow-y-auto flex-1 p-5 space-y-4">
          {/* 거래처 선택 */}
          <div>
            <label className="text-xs font-semibold text-neutral-500 mb-1 block">거래처 (딜러/아카데미)</label>
            <div className="relative">
              <input
                type="text"
                value={selectedCustomer ? selectedCustomer.name : customerQuery}
                onChange={(e) => {
                  if (selectedCustomer) { setSelectedCustomer(null); setCustomerName(''); }
                  setCustomerQuery(e.target.value);
                  setShowCustomerDropdown(true);
                }}
                onFocus={() => customerQuery.length >= 2 && setShowCustomerDropdown(true)}
                placeholder="거래처 검색 (2자 이상)"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
              />
              {selectedCustomer && (
                <button onClick={() => { setSelectedCustomer(null); setCustomerQuery(''); setCustomerName(''); }}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-neutral-400 hover:text-neutral-600">
                  <X size={14} />
                </button>
              )}
              {showCustomerDropdown && searchResults.length > 0 && !selectedCustomer && (
                <div className="absolute z-10 w-full mt-1 bg-white border border-neutral-200 rounded-lg shadow-lg max-h-48 overflow-y-auto">
                  {searchResults
                    .filter((c) => {
                      // 딜러/아카데미만 필터 (source로 판별)
                      const src = (c as unknown as Record<string, unknown>).source as string;
                      const type = (c as unknown as Record<string, unknown>).customer_type as string;
                      return type === 'dealer' || type === 'academy' || src === 'dealer' || src === 'academy';
                    })
                    .map((c) => (
                      <button
                        key={c.id}
                        onClick={() => {
                          setSelectedCustomer({
                            id: c.id,
                            name: c.name,
                            phone: c.phone,
                            customer_type: (c as unknown as Record<string, unknown>).customer_type as string,
                          });
                          setCustomerName(c.name);
                          setCustomerPhone(c.phone || '');
                          setShowCustomerDropdown(false);
                          // 장바구니 가격 재계산
                          const type = (c as unknown as Record<string, unknown>).customer_type as string;
                          setCart((prev) => prev.map((item) => {
                            if (!item.product_id) return item;
                            const prod = products.find((p) => p.id === item.product_id);
                            if (!prod) return item;
                            let price = prod.price;
                            if (type === 'dealer' && (prod as Record<string, unknown>).price_dealer) price = (prod as Record<string, unknown>).price_dealer as number;
                            if (type === 'academy' && (prod as Record<string, unknown>).price_academy) price = (prod as Record<string, unknown>).price_academy as number;
                            return { ...item, unit_price: price };
                          }));
                        }}
                        className="w-full flex items-center justify-between px-3 py-2 text-sm hover:bg-neutral-50 transition text-left"
                      >
                        <div>
                          <span className="font-medium">{c.name}</span>
                          {c.phone && <span className="text-xs text-neutral-400 ml-2">{formatPhone(c.phone)}</span>}
                        </div>
                        <span className={`text-[10px] px-1.5 py-0.5 rounded font-medium ${
                          (c as unknown as Record<string, unknown>).customer_type === 'dealer' ? 'bg-blue-100 text-blue-700' : 'bg-purple-100 text-purple-700'
                        }`}>
                          {(c as unknown as Record<string, unknown>).customer_type === 'dealer' ? '딜러' : '아카데미'}
                        </span>
                      </button>
                    ))}
                  {searchResults.filter((c) => {
                    const type = (c as unknown as Record<string, unknown>).customer_type as string;
                    return type === 'dealer' || type === 'academy';
                  }).length === 0 && (
                    <div className="px-3 py-2 text-xs text-neutral-400">딜러/아카데미 고객이 없습니다</div>
                  )}
                </div>
              )}
            </div>
            {/* 직접 입력 (검색 결과 없을 때) */}
            {!selectedCustomer && (
              <div className="grid grid-cols-2 gap-2 mt-2">
                <input type="text" value={customerName} onChange={(e) => setCustomerName(e.target.value)}
                  placeholder="거래처명 직접 입력"
                  className="h-8 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
                <input type="text" value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)}
                  placeholder="연락처"
                  className="h-8 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
              </div>
            )}
            {customerType === 'dealer' && <p className="text-xs text-blue-600 mt-1">딜러가 적용</p>}
            {customerType === 'academy' && <p className="text-xs text-purple-600 mt-1">아카데미가 적용</p>}
          </div>

          {/* 날짜 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">납품일</label>
              <input type="date" value={deliveryDate} onChange={(e) => setDeliveryDate(e.target.value)}
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">납품 예정일</label>
              <input type="date" value={expectedDate} onChange={(e) => setExpectedDate(e.target.value)}
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
          </div>

          {/* 제품 선택 */}
          <div>
            <label className="text-xs font-semibold text-neutral-500 mb-1 block">품목</label>
            <div className="relative">
              <input
                type="text"
                value={productSearch}
                onChange={(e) => setProductSearch(e.target.value)}
                placeholder="제품명 또는 SKU 검색"
                className="w-full h-8 px-3 rounded-lg border border-neutral-200 text-xs placeholder:text-neutral-400 focus:outline-none focus:ring-1 focus:ring-neutral-300"
              />
              {productSearch && (
                <div className="absolute z-20 w-full mt-1 max-h-[180px] overflow-y-auto bg-white border border-neutral-200 rounded-lg shadow-lg divide-y divide-neutral-50">
                  {filteredProducts.map((p) => {
                    const inCart = cart.find((c) => c.product_id === p.id);
                    const price = getPrice(p);
                    return (
                      <button
                        key={p.id}
                        onClick={() => { addProduct(p); setProductSearch(''); }}
                        className={`w-full flex items-center justify-between px-3 py-2 text-xs hover:bg-neutral-50 transition text-left ${inCart ? 'bg-neutral-900/5' : ''}`}
                      >
                        <span className="truncate font-medium">
                          {p.name}
                          {inCart && <span className="ml-1.5 text-neutral-900 font-bold">x{inCart.quantity}</span>}
                        </span>
                        <span className="text-neutral-400 shrink-0 ml-2">{formatKRW(price)}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
            {/* 임시 제품 직접 입력 */}
            <div className="flex gap-2 mt-2">
              <input type="text" value={customName} onChange={(e) => setCustomName(e.target.value)}
                placeholder="품목명 직접 입력" className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs placeholder:text-neutral-400" />
              <input type="number" value={customPrice} onChange={(e) => setCustomPrice(e.target.value)}
                placeholder="금액" className="w-24 h-7 px-2 rounded border border-neutral-200 text-xs text-right placeholder:text-neutral-400" />
              <button onClick={() => {
                if (!customName.trim() || !parseInt(customPrice)) return;
                setCart((prev) => [...prev, { product_name: customName.trim(), quantity: 1, unit_price: parseInt(customPrice) || 0 }]);
                setCustomName(''); setCustomPrice('');
              }} className="h-7 px-3 rounded bg-neutral-900 text-white text-[10px] font-semibold shrink-0">추가</button>
            </div>

            {/* 장바구니 */}
            {cart.length > 0 && (
              <div className="mt-2 space-y-1.5">
                {cart.map((item, idx) => (
                  <div key={idx} className="flex items-center gap-2 py-1 border-b border-neutral-50 last:border-0">
                    <div className="flex-1 min-w-0">
                      <p className="text-xs font-medium truncate">{item.product_name}</p>
                    </div>
                    <div className="flex items-center gap-1">
                      <button onClick={() => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, quantity: Math.max(1, c.quantity - 1) } : c))}
                        className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200 text-xs">-</button>
                      <input
                        type="number"
                        min={1}
                        value={item.quantity}
                        onChange={(e) => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, quantity: Math.max(1, parseInt(e.target.value) || 1) } : c))}
                        className="w-10 h-6 text-center text-xs font-bold border border-neutral-200 rounded bg-white focus:outline-none [appearance:textfield] [&::-webkit-inner-spin-button]:appearance-none"
                      />
                      <button onClick={() => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, quantity: c.quantity + 1 } : c))}
                        className="w-6 h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200 text-xs">+</button>
                    </div>
                    <input
                      type="number"
                      value={item.unit_price || ''}
                      onChange={(e) => setCart((prev) => prev.map((c, i) => i === idx ? { ...c, unit_price: parseInt(e.target.value) || 0 } : c))}
                      className="w-20 h-7 px-2 rounded border border-neutral-200 bg-warm-ivory text-xs text-right"
                    />
                    <span className="text-xs text-neutral-500 w-20 text-right">{formatKRW(item.quantity * item.unit_price)}</span>
                    <button onClick={() => setCart((prev) => prev.filter((_, i) => i !== idx))}
                      className="w-6 h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500 text-xs">x</button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* VAT + 증빙 + 결제 — 그루핑 */}
          <div className="grid grid-cols-3 gap-3">
            <div className="border border-neutral-200 rounded-lg p-2.5">
              <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">부가세</label>
              <div className="flex gap-1">
                {(['included', 'separate', 'none'] as const).map((v) => (
                  <button key={v} onClick={() => setVatType(v)}
                    className={`flex-1 py-1.5 rounded text-[10px] font-semibold transition ${
                      vatType === v ? 'bg-neutral-900 text-white' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                    }`}>
                    {{ included: '포함', separate: '별도', none: '미적용' }[v]}
                  </button>
                ))}
              </div>
            </div>
            <div className="border border-neutral-200 rounded-lg p-2.5">
              <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">증빙유형</label>
              <div className="flex gap-1">
                {(['expense_proof', 'tax_invoice', 'none'] as const).map((r) => (
                  <button key={r} onClick={() => setReceiptType(r)}
                    className={`flex-1 py-1.5 rounded text-[10px] font-semibold transition ${
                      receiptType === r ? 'bg-neutral-900 text-white' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                    }`}>
                    {RECEIPT_LABEL[r]}
                  </button>
                ))}
              </div>
            </div>
            <div className="border border-neutral-200 rounded-lg p-2.5">
              <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">결제상태</label>
              <div className="flex gap-1">
                {(['unpaid', 'partial', 'paid'] as const).map((ps) => (
                  <button key={ps} onClick={() => { setPaymentStatus(ps); if (ps !== 'partial') setPaidAmount(0); }}
                    className={`flex-1 py-1.5 rounded text-[10px] font-semibold transition ${
                      paymentStatus === ps
                        ? ps === 'paid' ? 'bg-green-600 text-white' : ps === 'partial' ? 'bg-yellow-500 text-white' : 'bg-red-500 text-white'
                        : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                    }`}>
                    {PAYMENT_LABEL[ps]}
                  </button>
                ))}
              </div>
              {paymentStatus === 'partial' && (
                <input type="number" value={paidAmount || ''} onChange={(e) => setPaidAmount(parseInt(e.target.value) || 0)}
                  placeholder="선납금 입력"
                  className="w-full h-7 px-2 mt-2 rounded border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-1 focus:ring-neutral-300" />
              )}
            </div>
          </div>

          {/* 결제수단 */}
          <div className="border border-neutral-200 rounded-lg p-2.5">
            <label className="text-[10px] font-semibold text-neutral-400 mb-1.5 block">결제수단</label>
            <div className="flex gap-1">
              {(['card', 'cash', 'transfer', 'mixed'] as const).map((m) => (
                <button key={m} onClick={() => setPaymentMethod(m)}
                  className={`flex-1 py-1.5 rounded text-xs font-semibold transition ${
                    paymentMethod === m ? 'bg-neutral-900 text-white' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'
                  }`}>
                  {{ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' }[m]}
                </button>
              ))}
            </div>
          </div>

          {/* 할인 + 메모 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">할인금액</label>
              <input type="number" value={discount || ''} onChange={(e) => setDiscount(parseInt(e.target.value) || 0)}
                placeholder="0"
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
            <div>
              <label className="text-xs font-semibold text-neutral-500 mb-1 block">메모</label>
              <input type="text" value={memo} onChange={(e) => setMemo(e.target.value)}
                placeholder="메모 (선택)"
                className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-warm-ivory text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300" />
            </div>
          </div>

          {/* 합계 */}
          {cart.length > 0 && (
            <div className="bg-neutral-50 rounded-lg p-3 space-y-1">
              <div className="flex justify-between text-xs text-neutral-500">
                <span>품목 합계</span><span>{formatKRW(itemTotal)}</span>
              </div>
              {discount > 0 && (
                <div className="flex justify-between text-xs text-red-500">
                  <span>할인</span><span>-{formatKRW(discount)}</span>
                </div>
              )}
              <div className="flex justify-between text-xs text-neutral-500">
                <span>공급가액</span><span>{formatKRW(supply)}</span>
              </div>
              {vatType !== 'none' && (
                <div className="flex justify-between text-xs text-neutral-500">
                  <span>부가세</span><span>{formatKRW(vat)}</span>
                </div>
              )}
              <div className="flex justify-between text-sm font-bold pt-1 border-t border-neutral-200">
                <span>합계</span><span>{formatKRW(totalAmount)}</span>
              </div>
            </div>
          )}
        </div>

        {/* 푸터 */}
        <div className="flex justify-end gap-2 px-5 py-3 border-t border-neutral-200">
          <Button variant="ghost" onClick={onClose}>취소</Button>
          <Button
            onClick={handleSubmit}
            disabled={cart.length === 0 || (!selectedCustomer && !customerName.trim()) || createDelivery.isPending}
          >
            {createDelivery.isPending ? '생성 중...' : '납품서 생성'}
          </Button>
        </div>
      </div>
    </div>
  );
}
