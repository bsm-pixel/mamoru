'use client';

import { useState, Suspense, useEffect, useRef } from 'react';
import { KeyHint } from '@/components/ui/key-hint';
import { MoneyInput, MoneyText } from '@/components/ui/money-input';
import { useUnsavedGuard } from '@/hooks/use-unsaved-guard';
import { useFormDraft, DraftBanner } from '@/hooks/use-form-draft';
import { useRouter, useSearchParams } from 'next/navigation';
import { useHotkeys, isDialogOpen } from '@/hooks/use-hotkeys';
import { groupArrowKeys } from '@/lib/ui/group-keys';
import Link from 'next/link';
import toast from 'react-hot-toast';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { CreateDeliveryModal } from '@/components/deliveries/create-delivery-modal';
import { useProducts, useCreateSale } from '@/hooks/use-sales';
import { formatKRW, calcVAT, toLocalDateString, channelFromConsultationType } from '@/lib/utils/format';
import { Minus, Plus, ScanLine, ShoppingBag, Trash2 } from 'lucide-react';
import { CustomerAutocomplete, type SelectedCustomer } from '@/components/shared/customer-autocomplete';
import { CustomerCreateModal } from '@/components/customers/customer-create-modal';
import { SerialPicker } from '@/components/sales/serial-picker';
import { resolveScan } from '@/lib/sales/resolve-scan';
import { useSerialConflictPrompt } from '@/components/sales/serial-conflict-dialog';
import { formatSerial, currentYear2, incrementSerial, normalizeSerial, isMSerial } from '@/lib/serial/format';
import { getUnitPrice, getProductDisplayName, hasGroupPrice } from '@/lib/utils/pricing';
import { usePriceGroups } from '@/hooks/use-price-groups';
import { useCustomerCatalog } from '@/hooks/use-customer-catalog';
import { TagSelector } from '@/components/shared/tag-selector';
import { useSetting } from '@/hooks/use-settings';
import type { Product } from '@/lib/supabase/types';
import { SkeletonRows } from '@/components/ui/skeleton';
import { matchSearch } from '@/lib/utils/search-match';
import { useTopProductIds } from '@/hooks/use-top-products';

interface CartItem {
  product: Product | null; // null이면 임시 제품
  customName?: string;     // 임시 제품명
  customPrice?: number;    // 임시 제품 가격
  // 2026-10-12: 직접 입력 품목의 분류. 'RS'=복원수리 / undefined=제품.
  //   전엔 직접 입력이 분류 없이 저장돼, 복원수리를 직접 입력하면 제품 매출로 잡혔다(집계 왜곡)
  customCategory?: string;
  quantity: number;
  unitPrice: number;
  selectedSerialIds: string[];
  manualSerials?: string[]; // 직접 입력 시리얼 번호
}

/** cart item의 고유 키 */
function cartItemKey(item: CartItem): string {
  return item.product ? item.product.id : `custom-${item.customName}-${item.customPrice}`;
}

const SERIAL_STATUS_LABEL: Record<string, string> = {
  in_stock: '재고', reserved: '예약', sold: '판매완료', returned: '반품', defective: '불량',
};

export default function NewSalePage() {
  return (
    <Suspense fallback={null}>
      <NewSaleRouter />
    </Suspense>
  );
}

/** 2026-05-26 Phase C: mode 분기 라우터 — ?mode=b2b 면 거래처 매출 입력 (모달), 그 외 기존 B2C 흐름 */
function NewSaleRouter() {
  const searchParams = useSearchParams();
  const mode = searchParams?.get('mode');
  if (mode === 'b2b') {
    return <B2BSaleEntry />;
  }
  return <NewSaleContent />;
}

/** B2B 거래처 매출 — CreateDeliveryModal 자동 표시 (풀스크린 모달 = 페이지 IA 등가) */
function B2BSaleEntry() {
  const router = useRouter();
  const searchParams = useSearchParams();
  // ?initial=repair 면 복원수리 탭으로 시작, 그 외 'delivery' (제품 납품)
  const initialMode = (searchParams?.get('initial') === 'repair' ? 'repair' : 'delivery') as 'delivery' | 'repair';

  return (
    <>
      <Topbar title="거래처 매출 입력" />
      <CreateDeliveryModal
        initialMode={initialMode}
        onClose={() => router.push('/sales')}
        onCreated={() => router.push('/sales')}
      />
    </>
  );
}

function NewSaleContent() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const contractId = searchParams?.get('contract_id') || null;
  const fromConsultationId = searchParams?.get('from_consultation') || null;
  const { data: products = [], isLoading: productsLoading } = useProducts();
  const createSale = useCreateSale();
  const priceGroups = usePriceGroups();
  const availableTags = useSetting<string[]>('customer.tags', []);

  const [saleDate, setSaleDate] = useState(toLocalDateString(new Date()));
  const [cart, setCart] = useState<CartItem[]>([]);
  // Phase A — 시리얼 다른 판매에서 이전 동의 플래그 (2026-05-18)
  const [allowSerialTransfer, setAllowSerialTransfer] = useState(false);
  const [selectedCustomer, setSelectedCustomer] = useState<SelectedCustomer | null>(null);
  // 복원수리 입력 (제품과 한 장바구니에 함께 — 2026-06-09 통합). deliveries 모달과 대칭.
  const [rep, setRep] = useState({ mamoruQty: 0, mamoruPrice: 10000, otherQty: 0, otherPrice: 20000 });
  const [repShipping, setRepShipping] = useState(false); // 복원수리 배송비 3,000원 포함 여부
  // 165: 가위를 맡아 감(출장 현장 수거) → 저장 시 복원수리 건 자동 생성 + 맡김 확인 알림톡. 끄면 매출 줄만 남는다(현장 즉시 수리 등)
  const [repTakeAway, setRepTakeAway] = useState(true);

  // 073/074: B2B 납품처 catalog (customer별 납품명 + 단가 자동 입력용)
  const { data: customerCatalogData } = useCustomerCatalog(selectedCustomer?.id);
  const catalogEntryMap = new Map(
    (customerCatalogData?.catalog || []).map((c) => [c.product_id, c])
  );
  const [customerName, setCustomerName] = useState('');
  const [customerPhone, setCustomerPhone] = useState('');
  const [paymentMethod, setPaymentMethod] = useState('card');
  const [paymentStatus, setPaymentStatus] = useState<'paid' | 'unpaid' | 'partial'>('paid');
  const [depositAmount, setDepositAmount] = useState(0);
  const [saleChannel, setSaleChannel] = useState('store'); // 기본 매장 (레거시 'offline'은 신규 입력에서 폐지)
  const [discount, setDiscount] = useState(0);
  const [memo, setMemo] = useState('');

  // 복합 결제 분리 금액
  const [mixedCard, setMixedCard] = useState(0);
  const [mixedCash, setMixedCash] = useState(0);
  const [mixedTransfer, setMixedTransfer] = useState(0);

  const [showCreateCustomer, setShowCreateCustomer] = useState(false);
  const [sourceConsultation, setSourceConsultation] = useState<{ id: string; unique_id: string } | null>(null);
  const customerType = selectedCustomer?.customer_type;

  // 070: 출장/매장상담 → 판매 link prefill
  // 077: 자동매칭된 기존 고객(이름이 다를 때) 안내 토스트
  useEffect(() => {
    if (!fromConsultationId) return;
    let cancelled = false;
    (async () => {
      try {
        const res = await fetch(`/api/consultation/${fromConsultationId}`);
        if (!res.ok) return;
        const { consultation } = await res.json();
        if (cancelled || !consultation) return;
        setSourceConsultation({ id: consultation.id, unique_id: consultation.unique_id });
        setCustomerName(consultation.name || '');
        setCustomerPhone(consultation.phone || '');
        // 상담 유형 → 판매 채널 자동설정 (매장방문→매장 / 출장요청→출장 / 톡상담→톡)
        const autoCh = channelFromConsultationType(consultation.consultation_type);
        if (autoCh) setSaleChannel(autoCh);
        if (consultation.customer_id) {
          // 077: 실제 customer 레코드 fetch — phone_normalized로 자동 매칭된 기존 고객의 정확한 이름 사용
          let customerName: string = consultation.name;
          try {
            const custRes = await fetch(`/api/customers/${consultation.customer_id}`);
            if (custRes.ok) {
              const { customer } = await custRes.json();
              if (customer?.name) {
                customerName = customer.name;
                // 이름이 다르면 자동매칭 안내 (전화번호 기준 SSOT)
                if (customer.name !== consultation.name) {
                  toast(
                    `동일 전화번호로 자동매칭됨\n접수명: ${consultation.name} → 기존 고객: ${customer.name}`,
                    { duration: 5000, icon: 'ℹ️' }
                  );
                }
              }
            }
          } catch {
            /* customer fetch 실패해도 consultation 데이터로 진행 */
          }
          setSelectedCustomer({
            id: consultation.customer_id,
            name: customerName,
            phone: consultation.phone,
            email: null,
            address_road: consultation.address_road || null,
            address_detail: consultation.address_detail || null,
            postcode: consultation.postcode || null,
            ecount_customer_code: null,
          } as SelectedCustomer);
        }
      } catch {
        /* 조용히 실패, 수동 입력 가능 */
      }
    })();
    return () => { cancelled = true; };
  }, [fromConsultationId]);
  const productTotal = cart.reduce((s, item) => s + item.unitPrice * item.quantity, 0);
  // 카트 전체에 이미 배정된 수동/자동 시리얼 — 품목 간 자동생성 중복 방지용 (2026-06-11 fix)
  const cartReservedSerials = cart.flatMap((it) => it.manualSerials || []);
  const REPAIR_SHIPPING_FEE = 3000;
  const repairTotal = rep.mamoruQty * rep.mamoruPrice + rep.otherQty * rep.otherPrice + (repShipping ? REPAIR_SHIPPING_FEE : 0);
  const hasRepair = rep.mamoruQty > 0 || rep.otherQty > 0;
  const totalAmount = productTotal + repairTotal; // 제품 + 복원수리 통합
  const finalAmount = totalAmount - discount;
  // 복합결제 = 받은 돈을 수단별로 쪼갠 것 → 합계가 곧 수납액. 결제상태/미수 자동 판정(회계 정합)
  const mixedTotal = mixedCard + mixedCash + mixedTransfer;
  const isMixedPay = paymentMethod === 'mixed';
  const paidAmount = isMixedPay ? Math.min(mixedTotal, finalAmount)
    : paymentStatus === 'paid' ? finalAmount
    : paymentStatus === 'unpaid' ? 0
    : Math.min(depositAmount, finalAmount);
  const effPaymentStatus = isMixedPay
    ? (mixedTotal >= finalAmount ? 'paid' : mixedTotal <= 0 ? 'unpaid' : 'partial')
    : paymentStatus;

  // 임시 제품 입력 상태
  const [customProductName, setCustomProductName] = useState('');
  const [customProductPrice, setCustomProductPrice] = useState('');
  const [customProductIsRepair, setCustomProductIsRepair] = useState(false);   // 직접 입력 분류(복원수리 여부)

  function addToCart(product: Product, serialId?: string) {
    // 074: catalog.unit_price 우선 사용 (해당 customer 등록된 맞춤가) → 없으면 group 단가 fallback
    const catalogEntry = catalogEntryMap.get(product.id);
    const customPrice = catalogEntry?.unit_price;
    const price = (customPrice && customPrice > 0)
      ? customPrice
      : getUnitPrice(product, customerType, priceGroups);
    setCart((prev) => {
      const existing = prev.find((item) => item.product?.id === product.id);
      if (existing) {
        return prev.map((item) =>
          item.product?.id === product.id
            ? { ...item, quantity: item.quantity + 1, selectedSerialIds: serialId ? [...(item.selectedSerialIds || []), serialId] : item.selectedSerialIds }
            : item
        );
      }
      return [...prev, { product, quantity: 1, unitPrice: price, selectedSerialIds: serialId ? [serialId] : [] }];
    });
  }

  // 바코드 스캔 → SKU=품목 추가 / 시리얼=제품+시리얼 배정(이미 판매·예약이면 차단)
  async function handleScan(code: string) {
    const res = await resolveScan(code, products);
    if (res.type === 'product') {
      addToCart(res.product);
      toast.success(`${res.product.name} 추가`);
    } else if (res.type === 'serial') {
      if (!res.product) { toast.error(`시리얼의 제품을 찾을 수 없습니다 (${res.serial.serial_number})`); return; }
      if (res.serial.status !== 'in_stock') {
        toast.error(`사용 불가 시리얼 — ${SERIAL_STATUS_LABEL[res.serial.status] || res.serial.status} 상태 (${res.serial.serial_number})`);
        return;
      }
      if (cart.some((it) => (it.selectedSerialIds || []).includes(res.serial.id))) {
        toast.error(`이미 담긴 시리얼입니다 (${res.serial.serial_number})`);
        return;
      }
      addToCart(res.product, res.serial.id);
      toast.success(`${res.product.name} · ${res.serial.serial_number} 추가`);
    } else if (res.type === 'serial-notfound') {
      toast.error(`해당 시리얼 없음: ${res.code}`);
    } else {
      toast.error(`바코드 매칭 제품 없음: ${res.code}`);
    }
  }

  function addCustomProduct() {
    const name = customProductName.trim();
    const price = parseInt(customProductPrice) || 0;
    if (!name || price <= 0) return;
    setCart((prev) => [
      ...prev,
      { product: null, customName: name, customPrice: price, quantity: 1, unitPrice: price, selectedSerialIds: [],
        customCategory: customProductIsRepair ? 'RS' : undefined },
    ]);
    setCustomProductName('');
    setCustomProductPrice('');
    setCustomProductIsRepair(false);
  }

  // 고객 변경 시 장바구니 가격 재계산 (등록 제품만)
  function recalcCartPrices(type?: string) {
    setCart((prev) => prev.map((item) => ({
      ...item,
      unitPrice: item.product ? getUnitPrice(item.product, type, priceGroups) : item.unitPrice,
    })));
  }

  function updateSerialIds(key: string, serialIds: string[]) {
    setCart((prev) =>
      prev.map((item) =>
        cartItemKey(item) === key ? { ...item, selectedSerialIds: serialIds } : item
      )
    );
  }

  function updateManualSerials(key: string, serials: string[]) {
    setCart((prev) =>
      prev.map((item) =>
        cartItemKey(item) === key ? { ...item, manualSerials: serials } : item
      )
    );
  }

  function updateQuantity(key: string, delta: number) {
    setCart((prev) =>
      prev
        .map((item) =>
          cartItemKey(item) === key
            ? { ...item, quantity: Math.max(0, item.quantity + delta) }
            : item
        )
        .filter((item) => item.quantity > 0)
    );
  }

  function removeFromCart(key: string) {
    setCart((prev) => prev.filter((item) => cartItemKey(item) !== key));
  }

  async function handleSubmit() {
    const name = selectedCustomer?.name || customerName.trim();
    if (!name) return;

    // 복원수리(마모루/타사/배송비)를 offline_sale_items category='RS' 로 저장 — 제품과 한 장바구니에 함께.
    //   product_name 고정 패턴 "복원수리 (마모루)" / "복원수리 (타사)" → 기존 includes('타사') 분류 기준 그대로 작동.
    const repairItems = [
      ...(rep.mamoruQty > 0 ? [{ product_name: '복원수리 (마모루)', category: 'RS', quantity: rep.mamoruQty, unit_price: rep.mamoruPrice, total_price: rep.mamoruQty * rep.mamoruPrice, serial_ids: [] as string[], manual_serials: [] as string[] }] : []),
      ...(rep.otherQty > 0 ? [{ product_name: '복원수리 (타사)', category: 'RS', quantity: rep.otherQty, unit_price: rep.otherPrice, total_price: rep.otherQty * rep.otherPrice, serial_ids: [] as string[], manual_serials: [] as string[] }] : []),
      ...(repShipping ? [{ product_name: '배송비', category: 'RS', quantity: 1, unit_price: REPAIR_SHIPPING_FEE, total_price: REPAIR_SHIPPING_FEE, serial_ids: [] as string[], manual_serials: [] as string[] }] : []),
    ];

    if (cart.length === 0 && !hasRepair) { toast.error('제품 또는 복원수리를 입력해주세요'); return; }
    // 165: 맡김 접수는 연락처가 있어야 한다(알림톡·출고 안내·고객 매칭)
    const phoneForRepair = selectedCustomer?.phone || customerPhone.trim();
    // 2026-11-09 출장 채널일 때만 수거 접수 (매장 손님이 두고 간 가위는 복원수리 접수에서 따로 받는다)
    const takeAway = hasRepair && repTakeAway && saleChannel === 'field';
    if (takeAway && !phoneForRepair) { toast.error('가위를 수거하려면 고객 연락처가 필요합니다'); return; }

    await createSale.mutateAsync({
      sale: {
        customer_id: selectedCustomer?.id || undefined,
        customer_name: name,
        customer_phone: selectedCustomer?.phone || customerPhone.trim() || undefined,
        total_amount: totalAmount,
        discount_amount: discount,
        paid_amount: paidAmount,
        payment_method: paymentMethod,
        payment_status: effPaymentStatus,
        payment_detail: paymentMethod === 'mixed'
          ? { card: mixedCard, cash: mixedCash, transfer: mixedTransfer }
          : { [paymentMethod]: paidAmount },
        sale_date: saleDate,
        sale_channel: saleChannel,
        customer_type: customerType || undefined,
        contract_id: contractId,
        source_consultation_id: sourceConsultation?.id || undefined,
        memo: memo.trim() || undefined,
      },
      items: [
        ...cart.map((item) => ({
          product_id: item.product?.id || undefined,
          // 073/074: catalog delivery_name 우선 사용 → 없으면 기존 fallback (price_groups display_name 또는 product.name)
          product_name: item.product
            ? ((catalogEntryMap.get(item.product.id)?.delivery_name?.trim()) || getProductDisplayName(item.product, customerType, priceGroups))
            : (item.customName || '임시 제품'),
          sku: item.product?.sku || undefined,
          category: item.product?.category || item.customCategory || undefined,
          quantity: item.quantity,
          unit_price: item.unitPrice,
          total_price: item.unitPrice * item.quantity,
          serial_ids: item.selectedSerialIds,
          manual_serials: item.manualSerials || [],
        })),
        ...repairItems, // 복원수리(category='RS') 라인 결합
      ],
      allow_serial_transfer: allowSerialTransfer, // Phase A — 사장님 명시 동의 시 서버 강탈 허용
      // 165: 가위를 맡아 감 → 서버가 판매와 연결된 복원수리 건(출장수거·결제는 이 판매)을 만든다
      ...(takeAway ? { field_repair: { qty_mamoru: rep.mamoruQty, qty_other: rep.otherQty } } : {}),
    });

    // 2026-11-15 다음 판매 입력 때 방금 고른 판매 채널 · 결제 수단이 먼저 골라져 있게 기억(복합 결제는 기억하지 않는다)
    try {
      localStorage.setItem('tms.sale.last', JSON.stringify({ channel: saleChannel, method: paymentMethod === 'mixed' ? 'card' : paymentMethod }));
    } catch { /* 저장 못 해도 무시 */ }
    draft.clear(); // 저장 성공 — 임시 저장분을 지운다
    router.push('/sales');
  }

  // 저장 가능 여부 — 하단 버튼 disabled 와 동일 기준. Ctrl+S 단축키와 공유
  const canSubmit = (!!selectedCustomer || !!customerName.trim()) && (cart.length > 0 || hasRepair) && !createSale.isPending;
  // Ctrl+S = 저장(등록). 폼 입력 중에도 동작(allowInInput) + 브라우저 저장창 차단(preventDefault)
  // 2026-11-15 Ctrl+Enter = 판매 등록(창의 "Ctrl+Enter = 저장"과 같은 키). 창이 떠 있으면 창에 양보한다.
  //   못 누르는 상태면 무엇이 빠졌는지 알려 준다(키보드만 쓸 때 버튼이 흐린 이유를 알 수 없어서).
  function submitByKey() {
    if (isDialogOpen() || createSale.isPending) return;
    if (!selectedCustomer && !customerName.trim()) { toast.error('고객을 먼저 골라 주세요 (F8)'); focusCustomer(); return; }
    if (cart.length === 0 && !hasRepair) { toast.error('제품을 먼저 담아 주세요'); productSearchRef.current?.focus(); return; }
    if (canSubmit) handleSubmit();
  }
  // F8 = 고객 칸으로. 이미 고른 고객이 있으면 그대로 두고 알려만 준다(실수로 지워지지 않게)
  function focusCustomer() {
    const el = document.getElementById('sale-customer-input') as HTMLInputElement | null;
    if (el) { el.focus(); el.select(); }
    else if (selectedCustomer) toast('고객이 이미 선택돼 있습니다 — 바꾸려면 이름 옆 X');
  }
  // F9 = 결제 수단 바꾸기 (카드 → 현금 → 이체 → 복합 → 카드). 버튼을 누른 것과 같은 길로 처리
  function cyclePayment() {
    const btns = Array.from(payGroupRef.current?.querySelectorAll<HTMLButtonElement>('button') || []);
    if (btns.length === 0) return;
    const i = btns.findIndex((b) => b.getAttribute('aria-checked') === 'true');
    btns[(i + 1) % btns.length].click();
  }
  useHotkeys([
    { combo: 'ctrl+s', allowInInput: true, handler: submitByKey },
    { combo: 'ctrl+enter', allowInInput: true, handler: submitByKey },
    { combo: 'f8', allowInInput: true, handler: () => { if (!isDialogOpen()) focusCustomer(); } },
    { combo: 'f9', allowInInput: true, handler: () => { if (!isDialogOpen()) cyclePayment(); } },
  ]);
  // 직전에 고른 판매 채널 · 결제 수단으로 시작 — 상담 · 계약서에서 넘어온 판매는 그쪽이 정한 값을 그대로 둔다
  useEffect(() => {
    if (fromConsultationId || contractId) return;
    const t = setTimeout(() => {
      try {
        const last = JSON.parse(localStorage.getItem('tms.sale.last') || 'null') as { channel?: string; method?: string } | null;
        if (last?.channel && ['store', 'field', 'talk', 'online'].includes(last.channel)) setSaleChannel(last.channel);
        if (last?.method && ['card', 'cash', 'transfer'].includes(last.method)) setPaymentMethod(last.method);
      } catch { /* 무시 */ }
    }, 0);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  const productSearchRef = useRef<HTMLInputElement>(null);
  const payGroupRef = useRef<HTMLDivElement>(null);
  // 화면이 열리면 커서를 제품 검색 칸에(바코드 스캐너도 이 칸으로 들어온다). 터치 기기는 자판이 튀어나오므로 제외
  useEffect(() => {
    if (typeof window !== 'undefined' && window.matchMedia?.('(pointer: coarse)').matches) return;
    productSearchRef.current?.focus();
  }, []);

  // 제품 검색/필터
  const [productSearch, setProductSearch] = useState('');
  const [productCategory, setProductCategory] = useState<string>('all');
  // 검색 결과 화살표 네비게이션 — ↑↓ 로 활성 행 지정, Enter 로 담기(결과 2개 이상일 때)
  const [activeIdx, setActiveIdx] = useState(-1);
  const activeRowRef = useRef<HTMLTableRowElement>(null);
  // 활성 행이 스크롤 밖이면 보이게 (setState 아님 → effect 안전)
  useEffect(() => { activeRowRef.current?.scrollIntoView({ block: 'nearest' }); }, [activeIdx]);

  const CATEGORY_LABEL: Record<string, string> = {
    BL: '블런트',
    TH: '틴닝',
    LO: '장가위',
    SL: '슬라이싱',
  };

  const matchedProducts = products.filter((p) => {
    if (p.category === 'RS') return false; // 복원수리는 상단 "복원수리" 모드로 입력 — 제품 선택 목록에서 제외
    const hit = !productSearch ||
      matchSearch(productSearch, p.name, p.sku); // 2026-11-15 초성 · 띄어쓰기/하이픈 무시
    const matchCat = productCategory === 'all' || p.category === productCategory;
    return hit && matchCat;
  });
  // 2026-11-15 검색도 분류도 안 골랐을 때는 제품 전체(160여 줄) 대신 「자주 파는 제품 10개」만 먼저 보여 준다
  const { data: topIds = [] } = useTopProductIds();
  const [showAllProducts, setShowAllProducts] = useState(false);
  const browsing = !productSearch && productCategory === 'all' && !showAllProducts;
  const frequentProducts = (() => {
    const byId = new Map(matchedProducts.map((p) => [p.id, p]));
    const top = topIds.map((id) => byId.get(id)).filter((p): p is Product => !!p);
    // 판매 기록이 아직 적으면 나머지를 목록 순서대로 채운다
    const rest = matchedProducts.filter((p) => !topIds.includes(p.id));
    return [...top, ...rest].slice(0, 10);
  })();
  const filteredProducts = browsing ? frequentProducts : matchedProducts;

  const submitDisabled = (!selectedCustomer && !customerName.trim()) || (cart.length === 0 && !hasRepair) || createSale.isPending;

  // 2026-11-15 넣다가 다른 메뉴를 누르면 한 번 묻는다 (저장 성공 뒤 스스로 이동할 때는 묻지 않음)
  const unsavedGuard = useUnsavedGuard(!createSale.isPending && !createSale.isSuccess && (cart.length > 0 || hasRepair || !!customerName.trim() || !!memo.trim()));

  // 2026-11-21 임시 저장 — 창을 닫거나 새로고침해도 이어서 쓰기. 상담 · 계약서에서 넘어온 입력(주소에 값이 붙은 경우)은 건드리지 않는다
  const draftSnapshot = { cart, selectedCustomer, rep, repShipping, repTakeAway, customerName, customerPhone, paymentMethod, paymentStatus, depositAmount, saleChannel, discount, memo, mixedCard, mixedCash, mixedTransfer };
  const draft = useFormDraft('sale-new', 1, draftSnapshot, cart.length > 0 || hasRepair || !!customerName.trim() || !!memo.trim(), !contractId && !fromConsultationId);
  const resumeDraft = () => {
    const d = draft.take();
    if (!d) return;
    setCart(d.cart); setSelectedCustomer(d.selectedCustomer); setRep(d.rep); setRepShipping(d.repShipping); setRepTakeAway(d.repTakeAway);
    setCustomerName(d.customerName); setCustomerPhone(d.customerPhone); setPaymentMethod(d.paymentMethod); setPaymentStatus(d.paymentStatus);
    setDepositAmount(d.depositAmount); setSaleChannel(d.saleChannel); setDiscount(d.discount); setMemo(d.memo);
    setMixedCard(d.mixedCard); setMixedCash(d.mixedCash); setMixedTransfer(d.mixedTransfer);
  };

  return (
    <>
      {unsavedGuard}
      <Topbar title="판매 입력" />

      {/* 2026-10-19 모바일: 하단 고정 바(장바구니·합계·판매 등록) 높이만큼 아래 여백 */}
      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4 pb-32 xl:pb-4 space-y-3">
        <DraftBanner at={draft.pendingAt} onResume={resumeDraft} onDiscard={draft.discard} />
        {/* 070: 상담 link 안내 배너 */}
        {sourceConsultation && (
          <div className="flex items-center gap-2 px-3 py-2 rounded-lg bg-blue-50 border border-blue-200 text-xs">
            <span className="text-blue-700 font-semibold">출장/매장상담에서 가져옴 ·</span>
            <Link href={`/consultations/${sourceConsultation.id}`} className="text-blue-700 font-mono font-bold underline">
              {sourceConsultation.unique_id}
            </Link>
            <span className="text-neutral-500 ml-auto">저장 시 이 판매가 상담에 연결되어 후기 관리가 자동 통합됩니다</span>
          </div>
        )}

        {/* 제목 + 검색 (그리드 밖, 좌열 너비만큼만) */}
        <div className="xl:w-[33.33%]">
            <h3 className="text-sm font-semibold text-neutral-700 mb-2">제품 선택</h3>
            {/* 2026-11-15 바코드 칸과 제품 검색 칸을 하나로 — 정확히 맞는 코드(SKU·바코드·시리얼)면 바로 담고, 아니면 이름 검색 */}
            {/* 검색 + 카테고리 필터 — 모바일은 카테고리를 검색 아래 줄로 (전엔 한 줄에 몰려 검색칸이 80px 로 줄었다) */}
            <div className="flex flex-col sm:flex-row sm:items-center gap-2">
              <div className="relative w-full sm:flex-1 min-w-0">
              <ScanLine size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400 pointer-events-none" />
              <input
                ref={productSearchRef}
                tabIndex={1}
                type="text"
                value={productSearch}
                onChange={(e) => { setProductSearch(e.target.value); setActiveIdx(-1); }}
                onKeyDown={(e) => {
                  // ↓/↑ : 결과 행 이동(활성 지정)
                  if (e.key === 'ArrowDown') { e.preventDefault(); setActiveIdx((i) => Math.min(i + 1, filteredProducts.length - 1)); return; }
                  if (e.key === 'ArrowUp') { e.preventDefault(); setActiveIdx((i) => Math.max(i - 1, 0)); return; }
                  if (e.key !== 'Enter' || e.ctrlKey || e.metaKey) return; // Ctrl+Enter 는 판매 등록 — 여기서 담지 않는다
                  const c = productSearch.trim();
                  // ① 코드 스캔(시리얼/정확 SKU·바코드) 우선 — 타이핑한 제품명은 무시
                  const exactSku = c ? products.find((p) => (p.sku || '').toLowerCase() === c.toLowerCase() || ((p.barcode || '') as string).toLowerCase() === c.toLowerCase()) : undefined;
                  if (c && (isMSerial(c) || exactSku)) { e.preventDefault(); handleScan(c); setProductSearch(''); setActiveIdx(-1); return; }
                  // ② ↑↓ 로 지정한 활성 행이 있으면 그 제품 담기
                  if (activeIdx >= 0 && activeIdx < filteredProducts.length) { e.preventDefault(); addToCart(filteredProducts[activeIdx]); setProductSearch(''); setActiveIdx(-1); return; }
                  // ③ 지정이 없으면 첫 번째 담기 (2026-11-15 검색 칸 규칙 통일 — 전엔 결과가 1개일 때만 담겼다)
                  if (c && filteredProducts.length >= 1) { e.preventDefault(); addToCart(filteredProducts[0]); setProductSearch(''); setActiveIdx(-1); return; }
                  // ④ 맞는 것이 하나도 없으면 알려 주고 글자를 통째로 선택 — 다음 스캔·입력이 그대로 덮어쓴다
                  if (c && filteredProducts.length === 0) { e.preventDefault(); toast.error(`맞는 제품이 없습니다: ${c}`); e.currentTarget.select(); }
                }}
                placeholder="제품 이름 · SKU 검색 또는 바코드 스캔"
                className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
              />
              </div>
              <div className="flex flex-wrap gap-1 sm:shrink-0">
                <button
                  onClick={() => { setProductCategory('all'); setActiveIdx(-1); }}
                  className={`h-8 px-2.5 text-xs rounded-md border transition ${
                    productCategory === 'all' ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                  }`}
                >전체</button>
                {Object.entries(CATEGORY_LABEL).map(([key, label]) => (
                  <button
                    key={key}
                    onClick={() => { setProductCategory(key); setActiveIdx(-1); }}
                    className={`h-8 px-2.5 text-xs rounded-md border transition ${
                      productCategory === key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                    }`}
                  >{label}</button>
                ))}
              </div>
            </div>
          </div>

        {/* 3열 그리드 — 검색바 아래 라인부터 동일 높이 시작 */}
        <div className="grid grid-cols-1 xl:grid-cols-3 gap-4 xl:items-start">
          {/* 좌: 제품 목록 + 복원수리 입력 (한 장바구니에 함께) */}
          <div className="space-y-3">
            <>
            {/* 제품 목록 테이블 */}
            {productsLoading ? (
              <SkeletonRows rows={4} />
            ) : (
              <Card padding={false}>
                <div className="flex items-center justify-between px-3 py-2 border-b border-neutral-100 text-xs">
                  <span className="font-semibold text-neutral-600">
                    {browsing ? '자주 파는 제품 10개' : `제품 ${filteredProducts.length}개`}
                  </span>
                  {!productSearch && productCategory === 'all' && (
                    <button type="button" tabIndex={-1} onClick={() => { setShowAllProducts((v) => !v); setActiveIdx(-1); }}
                      className="h-7 px-2.5 rounded-md border border-neutral-200 text-neutral-600 font-semibold hover:bg-neutral-50">
                      {showAllProducts ? '자주 파는 것만' : `전체 보기 (${matchedProducts.length})`}
                    </button>
                  )}
                </div>
                <div className="flex-1 overflow-y-auto" style={{ maxHeight: 'calc(100vh - 360px)', minHeight: '300px' }}>
                  <table className="w-full text-sm">
                    <thead className="bg-neutral-50 sticky top-0">
                      <tr className="text-xs text-neutral-500">
                        <th className="text-left px-3 py-2 font-medium">제품명</th>
                        <th className="text-right px-3 py-2 font-medium w-24">단가</th>
                        <th className="text-right px-3 py-2 font-medium w-14">재고</th>
                        <th className="text-center px-3 py-2 font-medium w-12">추가</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-neutral-100">
                      {filteredProducts.map((p, i) => {
                        const inCart = cart.find((c) => c.product?.id === p.id);
                        const price = getUnitPrice(p, customerType, priceGroups);
                        const isDiscounted = hasGroupPrice(p, customerType, priceGroups);
                        const isActive = i === activeIdx; // ↑↓ 로 지정된 행
                        return (
                          <tr
                            key={p.id}
                            ref={isActive ? activeRowRef : undefined}
                            onClick={() => addToCart(p)}
                            onMouseEnter={() => setActiveIdx(i)}
                            className={`cursor-pointer transition ${
                              isActive ? 'bg-stone-100 ring-1 ring-inset ring-neutral-900/60'
                                : inCart ? 'bg-neutral-900/5' : 'hover:bg-stone-50/60'
                            }`}
                          >
                            <td className="px-3 py-2.5">
                              <span className="font-medium text-stone-900">{getProductDisplayName(p, customerType, priceGroups)}</span>
                              {inCart && <span className="ml-2 text-xs font-bold text-neutral-900">×{inCart.quantity}</span>}
                            </td>
                            <td className="px-3 py-2.5 text-right">
                              <span className="font-bold">{formatKRW(price)}</span>
                              {isDiscounted && (
                                <span className="block text-[10px] text-neutral-400 line-through">{formatKRW(p.price)}</span>
                              )}
                            </td>
                            <td className="px-3 py-2.5 text-right text-xs text-neutral-500">{p.stock_quantity ?? '-'}</td>
                            <td className="px-3 py-2.5 text-center">
                              <button
                                tabIndex={-1} // 2026-11-15 제품마다 있는 버튼이라 Tab 에서 뺀다(키보드는 ↑↓ + Enter)
                                aria-label={`${p.name} 담기`}
                                onClick={(e) => { e.stopPropagation(); addToCart(p); }}
                                className="w-7 h-7 rounded-md bg-neutral-100 hover:bg-neutral-200 text-neutral-600 text-sm font-bold transition"
                              >+</button>
                            </td>
                          </tr>
                        );
                      })}
                      {filteredProducts.length === 0 && (
                        <tr>
                          <td colSpan={4} className="px-3 py-6 text-center text-neutral-400 text-xs">
                            {productSearch ? '검색 결과가 없습니다' : '등록된 제품이 없습니다'}
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </Card>
            )}

            {/* 임시 제품 직접 입력 */}
            <Card>
              <h3 className="text-sm font-semibold text-neutral-700 mb-2">+ 직접 입력</h3>
              <p className="text-xs text-neutral-400 mb-2">등록되지 않은 제품을 직접 입력합니다 (빗, 소모품 등)</p>
              {/* 분류 — 복원수리를 직접 입력할 때 제품 매출로 잡히지 않게 (2026-10-12) */}
              <div className="flex gap-1.5 mb-2">
                {([[false, '제품'], [true, '복원수리']] as const).map(([v, label]) => (
                  <button
                    key={label}
                    type="button"
                    onClick={() => setCustomProductIsRepair(v)}
                    className={`px-2.5 py-1 rounded-lg border text-xs font-semibold transition ${
                      customProductIsRepair === v ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-600 border-neutral-200 hover:bg-stone-50'
                    }`}
                  >
                    {label}
                  </button>
                ))}
                <span className="self-center text-[11px] text-neutral-400">
                  {customProductIsRepair ? '복원수리 매출로 집계됩니다' : '제품 매출로 집계됩니다'}
                </span>
              </div>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={customProductName}
                  onChange={(e) => setCustomProductName(e.target.value)}
                  placeholder="제품명"
                  className="flex-1 min-w-0 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
                />
                <MoneyText
                  value={customProductPrice}
                  onChange={(val) => setCustomProductPrice(val)}
                  placeholder="금액"
                  className="w-24 sm:w-28 shrink-0 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-neutral-300"
                />
                <Button
                  size="sm"
                  variant="secondary"
                  onClick={addCustomProduct}
                  disabled={!customProductName.trim() || !(parseInt(customProductPrice) > 0)}
                >
                  추가
                </Button>
              </div>
            </Card>
            </>
            {/* 복원수리 입력 — 제품과 한 장바구니에 함께 (2026-06-09 통합) */}
            <Card>
              <h3 className="text-sm font-semibold text-stone-900 mb-1">복원수리 입력 (선택)</h3>
              <p className="text-xs text-neutral-400 mb-3">마모루 / 타사 가위를 자루 수와 단가로 입력합니다. 제품과 함께 등록됩니다.</p>
              <div className="space-y-3">
                {([
                  { key: 'mamoru' as const, label: '마모루 가위', qty: rep.mamoruQty, price: rep.mamoruPrice, basePrice: '10,000' },
                  { key: 'other' as const, label: '타사 가위', qty: rep.otherQty, price: rep.otherPrice, basePrice: '20,000' },
                ]).map((r) => (
                  <div key={r.key} className="p-3 rounded-lg border border-neutral-200 bg-stone-50/40 space-y-2">
                    <p className="text-xs font-semibold text-neutral-700">{r.label}</p>
                    <div className="grid grid-cols-2 gap-2">
                      <div>
                        <label className="text-[10px] text-neutral-500 block mb-0.5">수량 (자루)</label>
                        <input type="number" min={0} value={r.qty || ''} placeholder="0"
                          onChange={(e) => { const v = Math.max(0, parseInt(e.target.value) || 0); setRep((p) => r.key === 'mamoru' ? { ...p, mamoruQty: v } : { ...p, otherQty: v }); }}
                          className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-neutral-300" />
                      </div>
                      <div>
                        <label className="text-[10px] text-neutral-500 block mb-0.5">단가 (원)</label>
                        <MoneyInput value={r.price} showZero
                          onChange={(val) => { const v = val; setRep((p) => r.key === 'mamoru' ? { ...p, mamoruPrice: v } : { ...p, otherPrice: v }); }}
                          className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-neutral-300" />
                      </div>
                    </div>
                    <p className="text-[10px] text-neutral-400">기본 {r.basePrice}원 · 수정 가능 · 소계 {formatKRW(r.qty * r.price)}</p>
                  </div>
                ))}
                {/* 배송비 (택배 발송 시) */}
                <button type="button" onClick={() => setRepShipping((v) => !v)}
                  className={`w-full flex items-center justify-between px-3 py-2 rounded-lg border text-xs font-medium transition ${repShipping ? 'border-neutral-900 bg-neutral-900 text-white' : 'border-neutral-200 bg-white text-neutral-600 hover:bg-neutral-50'}`}>
                  <span>{repShipping ? '✓ ' : '+ '}배송비 {formatKRW(REPAIR_SHIPPING_FEE)} 포함</span>
                  {repShipping && <span>{formatKRW(REPAIR_SHIPPING_FEE)}</span>}
                </button>
                <div className="flex justify-between text-sm font-bold pt-1 border-t border-neutral-100">
                  <span>복원수리 소계 ({rep.mamoruQty + rep.otherQty}자루{repShipping ? ' + 배송비' : ''})</span>
                  <span className="text-stone-900">{formatKRW(repairTotal)}</span>
                </div>
                {/* 165: 출장 현장에서 가위를 맡아 감 → 복원수리 접수 자동 생성 (결제는 이 판매로 끝) */}
                {hasRepair && saleChannel === 'field' && (
                  <label className="flex items-start gap-2 p-3 rounded-lg border border-neutral-200 bg-white cursor-pointer">
                    <input type="checkbox" checked={repTakeAway} onChange={(e) => setRepTakeAway(e.target.checked)}
                      className="mt-0.5 w-4 h-4 accent-neutral-900 shrink-0" />
                    <span className="min-w-0">
                      {/* 2026-11-09 문구: '가위를 맡아 가요' → 사장님 관점 '수거했어요' (고객 알림톡 문구는 고객 관점 그대로) */}
                      <span className="block text-xs font-semibold text-neutral-800">고객 가위를 수거했어요 — 복원수리 접수 함께 만들기</span>
                      <span className="block text-[11px] text-neutral-500 leading-relaxed mt-0.5">
                        {repTakeAway
                          ? '저장하면 복원수리 「진행중」에 출장수거 건이 바로 생기고(입고 확인 불필요), 고객에게 수거 확인 알림톡이 갑니다. 결제는 이 판매로 끝나 비용안내·입금확인은 건너뜁니다.'
                          : '현장에서 바로 수리해 건넨 경우 등 — 매출만 기록하고 복원수리 접수는 만들지 않습니다.'}
                      </span>
                    </span>
                  </label>
                )}
                {hasRepair && saleChannel !== 'field' && (
                  <p className="text-[11px] text-neutral-400 leading-relaxed break-keep">
                    매출만 기록됩니다. 출장에서 가위를 수거했다면 판매 채널을 「출장」으로 고르세요 — 복원수리 접수가 함께 만들어집니다.
                  </p>
                )}
              </div>
            </Card>
          </div>

          {/* 중: 고객 + 결제 */}
          <div className="space-y-3">
            {/* 고객 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-stone-900 mb-3">고객 정보</h3>
              <CustomerAutocomplete
                inputId="sale-customer-input"
                inputTabIndex={2}
                selectedCustomer={selectedCustomer}
                onSelect={(c) => {
                  setSelectedCustomer(c); setCustomerName(c.name); setCustomerPhone(c.phone || ''); recalcCartPrices(c.customer_type);
                  // 고르면 검색 칸이 사라진다 → 커서를 다음 칸(결제 수단)으로 넘겨 Tab 흐름이 끊기지 않게
                  requestAnimationFrame(() => payGroupRef.current?.querySelector<HTMLButtonElement>('button[aria-checked="true"]')?.focus());
                }}
                onClear={() => { setSelectedCustomer(null); setCustomerName(''); setCustomerPhone(''); recalcCartPrices(undefined); }}
              />
              {!selectedCustomer && (
                <button type="button" onClick={() => setShowCreateCustomer(true)}
                  className="w-full mt-2 py-2 rounded-lg border border-dashed border-neutral-300 text-xs text-neutral-500 hover:bg-neutral-50 transition">
                  + 신규 고객 등록
                </button>
              )}
              {selectedCustomer && (
                <div className="mt-3 pt-3 border-t border-neutral-100 space-y-1.5 text-xs text-neutral-500">
                  {selectedCustomer.address_road && (
                    <div>
                      <span className="text-neutral-400">주소: </span>
                      <span>{selectedCustomer.postcode || ''} {selectedCustomer.address_road || ''} {selectedCustomer.address_detail || ''}</span>
                    </div>
                  )}
                  {/* 태그 — 편집 가능 */}
                  {availableTags.length > 0 && (
                    <div>
                      <span className="text-neutral-400 block mb-1">태그</span>
                      <TagSelector
                        availableTags={availableTags}
                        selectedTags={selectedCustomer.tags || []}
                        onChange={async (tags) => {
                          setSelectedCustomer({ ...selectedCustomer, tags });
                          // 즉시 DB 반영
                          try {
                            await fetch(`/api/customers/${selectedCustomer.id}`, {
                              method: 'PATCH',
                              headers: { 'Content-Type': 'application/json' },
                              body: JSON.stringify({ tags }),
                            });
                          } catch { /* 조용히 실패 */ }
                        }}
                      />
                    </div>
                  )}
                  {/* 메모 — 고정 높이 영역 (메모 없어도 공간 유지) */}
                  <div className="min-h-[48px] p-2 rounded-lg bg-neutral-50 border border-neutral-100">
                    {selectedCustomer.memo ? (
                      <p className="text-xs text-neutral-600">{selectedCustomer.memo}</p>
                    ) : (
                      <p className="text-xs text-neutral-300 italic">고객 메모 없음</p>
                    )}
                  </div>
                </div>
              )}
              {customerType === 'dealer' && <p className="text-xs text-purple-600 mt-1">딜러가 적용 중</p>}
              {customerType === 'academy' && <p className="text-xs text-emerald-600 mt-1">아카데미가 적용 중</p>}
            </Card>

            <CustomerCreateModal
              open={showCreateCustomer}
              onClose={() => setShowCreateCustomer(false)}
              onCreated={(c) => { setSelectedCustomer({ id: c.id, name: c.name, phone: c.phone, customer_type: c.customer_type } as SelectedCustomer); setCustomerName(c.name); setCustomerPhone(c.phone); recalcCartPrices(c.customer_type); }}
            />

            {/* 결제 정보 */}
            <Card>
              <h3 className="text-sm font-semibold text-stone-900 mb-3">결제 정보</h3>
              <div className="space-y-3">
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">판매일</label>
                  <input type="date" value={saleDate} max={toLocalDateString(new Date())} onChange={(e) => setSaleDate(e.target.value)}
                    className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-neutral-300" />
                </div>
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">판매 채널</label>
                  <div className="flex gap-1" role="radiogroup" aria-label="판매 채널" onKeyDown={groupArrowKeys}>
                    {([
                      { value: 'store', label: '매장', color: 'bg-neutral-800 text-white' },
                      { value: 'field', label: '출장', color: 'bg-emerald-600 text-white' },
                      { value: 'talk', label: '톡', color: 'bg-yellow-500 text-white' },
                      { value: 'online', label: '온라인(아임웹)', color: 'bg-blue-600 text-white' },
                    ] as const).map((ch) => (
                      <button key={ch.value} type="button" role="radio" aria-checked={saleChannel === ch.value} tabIndex={saleChannel === ch.value ? 0 : -1} onClick={() => setSaleChannel(ch.value)}
                        className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition ${saleChannel === ch.value ? ch.color : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>{ch.label}</button>
                    ))}
                  </div>
                </div>
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">결제 수단</label>
                  <div ref={payGroupRef} className="flex gap-1" role="radiogroup" aria-label="결제 수단 (F9)" onKeyDown={groupArrowKeys}>
                    {(['card', 'cash', 'transfer', 'mixed'] as const).map((method) => (
                      <button key={method} type="button" role="radio" aria-checked={paymentMethod === method} tabIndex={paymentMethod === method ? 3 : -1} onClick={() => { setPaymentMethod(method); if (method === 'mixed') { setMixedCard(0); setMixedCash(0); setMixedTransfer(0); } }}
                        className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition ${paymentMethod === method ? 'bg-stone-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
                        {{ card: '카드', cash: '현금', transfer: '이체', mixed: '복합' }[method]}
                      </button>
                    ))}
                  </div>
                  {paymentMethod === 'mixed' && (
                    <div className="mt-2 p-2 rounded-lg bg-neutral-50 border border-neutral-200 space-y-1.5">
                      {(['카드', '현금', '이체'] as const).map((label, i) => {
                        const val = [mixedCard, mixedCash, mixedTransfer][i];
                        const setter = [setMixedCard, setMixedCash, setMixedTransfer][i];
                        return (
                          <div key={label} className="flex items-center gap-1.5">
                            <span className="text-[10px] text-neutral-500 w-8">{label}</span>
                            <MoneyInput value={val} onChange={setter}
                              className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs text-right bg-white focus:outline-none focus:ring-1 focus:ring-stone-400" placeholder="0" />
                          </div>
                        );
                      })}
                      {(() => { const t = mixedCard + mixedCash + mixedTransfer; const short = finalAmount - t; return (
                        <div className={`flex justify-between text-[10px] font-semibold pt-1 border-t border-neutral-200 ${short === 0 ? 'text-green-600' : short > 0 ? 'text-red-500' : 'text-orange-500'}`}>
                          <span>합계 (= 받은 금액)</span><span>{formatKRW(t)}{short === 0 ? ' ✓ 완납' : short > 0 ? ` · 미수 ${formatKRW(short)}` : ` · 초과 ${formatKRW(-short)}`}</span>
                        </div>
                      ); })()}
                    </div>
                  )}
                </div>
                {/* 복합결제는 합계가 결제상태를 자동 결정하므로 수동 결제상태 UI 숨김 */}
                {!isMixedPay && (
                <div>
                  <label className="text-xs text-neutral-500 mb-1 block">결제 상태</label>
                  <div className="flex gap-1" role="radiogroup" aria-label="결제 상태" onKeyDown={groupArrowKeys}>
                    {([{ value: 'paid' as const, label: '결제완료', color: 'bg-green-600 text-white' }, { value: 'partial' as const, label: '부분결제', color: 'bg-yellow-500 text-white' }, { value: 'unpaid' as const, label: '미결제', color: 'bg-red-500 text-white' }]).map((ps) => (
                      <button key={ps.value} type="button" role="radio" aria-checked={paymentStatus === ps.value} tabIndex={paymentStatus === ps.value ? 0 : -1} onClick={() => setPaymentStatus(ps.value)}
                        className={`flex-1 py-1.5 rounded-lg text-xs font-semibold transition ${paymentStatus === ps.value ? ps.color : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>{ps.label}</button>
                    ))}
                  </div>
                  {paymentStatus === 'partial' && (
                    <div className="mt-1.5">
                      <MoneyInput value={depositAmount} onChange={setDepositAmount} placeholder="입금액"
                        className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                    </div>
                  )}
                  {paymentStatus !== 'paid' && finalAmount > 0 && <p className="text-[10px] text-red-500 mt-1 font-medium">미수금: {formatKRW(finalAmount - paidAmount)}</p>}
                </div>
                )}
                <div>
                  <label className="text-xs text-neutral-500">할인 금액</label>
                  <MoneyInput value={discount} onChange={setDiscount} placeholder="0"
                    className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">메모</label>
                  <input type="text" tabIndex={4} value={memo} onChange={(e) => setMemo(e.target.value)} placeholder="메모 (선택)"
                    className="w-full h-8 px-2 rounded-lg border border-neutral-200 bg-stone-50 text-sm focus:outline-none focus:ring-2 focus:ring-stone-400" />
                </div>
              </div>
            </Card>
          </div>

          {/* 우: 장바구니 + 판매등록 */}
          {/* 2026-11-15 PC: 제품 목록을 내려도 장바구니 · [판매 등록]이 화면에 남는다 */}
          <div className="space-y-3 xl:sticky xl:top-[72px] xl:self-start xl:max-h-[calc(100vh-88px)] xl:overflow-y-auto" id="mm-cart">
            <Card>
              <h3 className="text-sm font-semibold text-stone-900 mb-3 flex items-center gap-2">
                <ShoppingBag size={16} />
                장바구니{cart.length > 0 ? ` (${cart.length})` : ''}{hasRepair || repShipping ? ' + 복원수리' : ''}
              </h3>
              {cart.length === 0 && !hasRepair && !repShipping ? (
                <p className="text-xs text-neutral-400 text-center py-4">제품 또는 복원수리를 입력해주세요</p>
              ) : (
                <div className="space-y-2">
                  {cart.map((item) => {
                    const key = cartItemKey(item);
                    const name = item.product ? getProductDisplayName(item.product, customerType) : (item.customName || '임시 제품');
                    const isCustom = !item.product;
                    return (
                      <div key={key}>
                        <div className="flex items-center gap-2">
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-1.5">
                              <p className="text-sm font-medium truncate">{name}</p>
                              {isCustom && <span className="text-[10px] px-1.5 py-0.5 rounded bg-orange-100 text-orange-600 font-medium shrink-0">임시</span>}
                            </div>
                            <p className="text-xs text-neutral-500">{formatKRW(item.unitPrice)} x {item.quantity} = {formatKRW(item.unitPrice * item.quantity)}</p>
                          </div>
                          <div className="flex items-center gap-1">
                            <button title="수량 줄이기" aria-label="수량 줄이기" onClick={() => updateQuantity(key, -1)} className="w-8 h-8 xl:w-6 xl:h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Minus size={12} /></button>
                            <input type="number" min="1" value={item.quantity}
                              onChange={(e) => { const val = parseInt(e.target.value) || 1; setCart((prev) => prev.map((it) => cartItemKey(it) === key ? { ...it, quantity: Math.max(1, val) } : it)); }}
                              className="w-10 h-8 xl:h-6 text-center text-sm font-semibold border border-neutral-200 rounded bg-white focus:outline-none focus:ring-1 focus:ring-neutral-300 [appearance:textfield] [&::-webkit-inner-spin-button]:appearance-none" />
                            <button title="수량 늘리기" aria-label="수량 늘리기" onClick={() => updateQuantity(key, 1)} className="w-8 h-8 xl:w-6 xl:h-6 rounded bg-neutral-100 flex items-center justify-center hover:bg-neutral-200"><Plus size={12} /></button>
                            <button title="삭제" aria-label="삭제" onClick={() => removeFromCart(key)} className="w-8 h-8 xl:w-6 xl:h-6 rounded bg-red-50 flex items-center justify-center hover:bg-red-100 text-red-500"><Trash2 size={12} /></button>
                          </div>
                        </div>
                        {item.product && customerType !== 'dealer' && customerType !== 'academy' ? (
                          <SerialPicker productId={item.product.id} quantity={item.quantity} selectedSerialIds={item.selectedSerialIds}
                            onSelect={(ids) => updateSerialIds(key, ids)} manualSerials={item.manualSerials || []} onManualSerialsChange={(s) => updateManualSerials(key, s)}
                            onTransferConsent={() => setAllowSerialTransfer(true)}
                            reservedSerials={cartReservedSerials.filter((sn) => !(item.manualSerials || []).includes(sn))} />
                        ) : item.product && (customerType === 'dealer' || customerType === 'academy') ? (
                          <p className="text-[10px] text-neutral-400 mt-1">보관창고에서 출고 (시리얼 미부여)</p>
                        ) : !item.product ? (
                          <ManualSerialInput serials={item.manualSerials || []} onChange={(s) => updateManualSerials(key, s)}
                            onTransferConsent={() => setAllowSerialTransfer(true)}
                            reservedSerials={cartReservedSerials.filter((sn) => !(item.manualSerials || []).includes(sn))} />
                        ) : null}
                      </div>
                    );
                  })}
                  {/* 복원수리 라인 (왼쪽 복원수리 입력 반영) */}
                  {rep.mamoruQty > 0 && <div className="flex items-center justify-between text-sm py-1"><span className="font-medium">복원수리 (마모루) × {rep.mamoruQty}자루</span><span className="text-neutral-600">{formatKRW(rep.mamoruQty * rep.mamoruPrice)}</span></div>}
                  {rep.otherQty > 0 && <div className="flex items-center justify-between text-sm py-1"><span className="font-medium">복원수리 (타사) × {rep.otherQty}자루</span><span className="text-neutral-600">{formatKRW(rep.otherQty * rep.otherPrice)}</span></div>}
                  {repShipping && <div className="flex items-center justify-between text-sm py-1"><span className="font-medium">배송비</span><span className="text-neutral-600">{formatKRW(REPAIR_SHIPPING_FEE)}</span></div>}
                </div>
              )}
              {(cart.length > 0 || repairTotal > 0) && (
                <div className="mt-3 pt-3 border-t border-neutral-100 space-y-1">
                  <div className="flex justify-between text-sm"><span className="text-neutral-500">소계</span><span className="font-semibold">{formatKRW(totalAmount)}</span></div>
                  {discount > 0 && <div className="flex justify-between text-sm"><span className="text-neutral-500">할인</span><span className="text-red-600">-{formatKRW(discount)}</span></div>}
                  <div className="flex justify-between text-sm font-bold"><span>결제 금액</span><span className="text-stone-900">{formatKRW(paidAmount)}</span></div>
                  {(() => {
                    const cardAmt = paymentMethod === 'card' ? paidAmount : paymentMethod === 'mixed' ? mixedCard : 0;
                    if (cardAmt <= 0) return null;
                    const { supply, vat } = calcVAT(cardAmt);
                    return (
                      <div className="mt-2 pt-2 border-t border-dashed border-neutral-200 space-y-0.5">
                        <div className="flex justify-between text-xs text-neutral-500"><span>카드 공급가액</span><span>{formatKRW(supply)}</span></div>
                        <div className="flex justify-between text-xs text-neutral-500"><span>부가세 (10%)</span><span>{formatKRW(vat)}</span></div>
                        {paymentMethod === 'mixed' && <div className="flex justify-between text-xs text-neutral-400"><span>현금/이체</span><span>{formatKRW(mixedCash + mixedTransfer)}</span></div>}
                      </div>
                    );
                  })()}
                </div>
              )}
            </Card>

            <div className="hidden xl:block">
              <Button className="w-full" tabIndex={5} title="Ctrl+Enter"
                disabled={submitDisabled}
                onClick={handleSubmit}>
                {createSale.isPending ? '등록 중...' : `판매 등록 (${formatKRW(paidAmount)})`}
                <KeyHint>Ctrl Enter</KeyHint>
              </Button>
            </div>
          </div>
        </div>
      </div>

      {/* 2026-10-19 모바일 하단 고정 바 — 상품을 고르는 동안에도 장바구니·합계·[판매 등록]이 보이게
          (전엔 맨 아래까지 스크롤해야 등록 버튼이 있었다). 왼쪽을 누르면 장바구니로 이동 */}
      <div className="xl:hidden fixed left-3 right-3 z-30 bottom-[calc(3.5rem+env(safe-area-inset-bottom)+8px)] md:bottom-4 rounded-2xl bg-stone-900 text-white shadow-xl px-3 py-2.5 flex items-center gap-3">
        <button type="button" onClick={() => document.getElementById('mm-cart')?.scrollIntoView({ behavior: 'smooth', block: 'start' })} className="flex-1 min-w-0 text-left">
          <p className="text-[11px] text-white/60">
            장바구니 {cart.reduce((n, it) => n + it.quantity, 0)}개{hasRepair || repShipping ? ' + 복원수리' : ''}
          </p>
          <p className="text-base font-bold tabular-nums truncate">{formatKRW(paidAmount)}</p>
        </button>
        <button type="button" tabIndex={5} disabled={submitDisabled} onClick={handleSubmit}
          className="h-11 px-5 rounded-xl bg-white text-stone-900 text-sm font-bold disabled:opacity-40 whitespace-nowrap">
          {createSale.isPending ? '등록 중...' : '판매 등록'}
        </button>
      </div>
    </>
  );
}

/** 직접입력 품목용 시리얼 직접 입력 */
function ManualSerialInput({ serials, onChange, onTransferConsent, reservedSerials = [] }: { serials: string[]; onChange: (s: string[]) => void; onTransferConsent?: () => void; reservedSerials?: string[] }) {
  const [input, setInput] = useState('');
  const [open, setOpen] = useState(false);
  const { prompt: promptConflict, dialog: conflictDialog } = useSerialConflictPrompt();

  // Phase A — 시리얼 추가 전 다른 판매와 중복 검증 (모달 다이얼로그) (2026-05-18)
  async function confirmIfDuplicate(serial: string): Promise<boolean> {
    try {
      const res = await fetch(`/api/serials/check-duplicate?serial=${encodeURIComponent(serial)}`);
      const data = await res.json();
      if (!data.exists) return true;
      const ok = await promptConflict({
        serial,
        sale_number: data.sale_number,
        customer_name: data.customer_name,
        product_name: data.product_name,
        sale_date: data.sale_date,
        status: data.status,
      });
      if (ok) onTransferConsent?.();
      return ok;
    } catch {
      return true; // API 실패 시 보수적 허용
    }
  }

  async function addSerial(value: string) {
    const trimmed = value.trim();
    if (!trimmed) return;
    const ok = await confirmIfDuplicate(trimmed);
    if (!ok) return;
    onChange([...serials, trimmed]);
    setInput('');
  }

  return (
    <div className="mt-1">
      {conflictDialog}
      <button type="button" onClick={() => setOpen(!open)}
        className={`text-xs px-2 py-1 rounded transition ${serials.length > 0 ? 'bg-blue-50 text-blue-700' : 'bg-neutral-50 text-neutral-500 hover:bg-neutral-100'}`}>
        # 시리얼 {serials.length > 0 ? `${serials.length}개` : '직접 입력'}
      </button>
      {open && (
        <div className="mt-1 p-2 rounded border border-neutral-200 bg-white space-y-1.5">
          {serials.map((s, i) => (
            <div key={i} className="flex items-center gap-1.5">
              <span className="font-mono text-xs text-neutral-700 flex-1">{s}</span>
              <button type="button" onClick={() => onChange(serials.filter((_, j) => j !== i))} className="text-red-400 hover:text-red-600 text-xs">×</button>
            </div>
          ))}
          <div className="flex gap-1.5">
            <input type="text" value={input} onChange={(e) => setInput(e.target.value)}
              onKeyDown={(e) => { if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && input.trim()) { e.preventDefault(); addSerial(input); } }}
              placeholder="시리얼 번호 입력 후 엔터"
              className="flex-1 h-7 px-2 rounded border border-neutral-200 text-xs font-mono placeholder:text-neutral-400" />
            <button type="button" onClick={() => addSerial(input)}
              className="px-2 py-1 text-xs bg-neutral-900 text-white rounded">추가</button>
          </div>
          <button type="button" onClick={async () => {
            try {
              const res = await fetch('/api/serials/batch');
              const data = await res.json();
              let next: string = data.next_serial || formatSerial(currentYear2(), 1);
              // 이 품목 + 카트 내 다른 품목 배정 번호 모두 회피 (품목 간 중복 방지) 2026-06-11
              const taken = new Set([...serials, ...reservedSerials].map(normalizeSerial));
              while (taken.has(normalizeSerial(next))) next = incrementSerial(next);
              const ok = await confirmIfDuplicate(next);
              if (!ok) return;
              onChange([...serials, next]);
            } catch { /* ignore */ }
          }} className="text-[10px] text-green-600 hover:text-green-800 font-medium">
            자동 번호 생성
          </button>
        </div>
      )}
      {!open && serials.length > 0 && (
        <div className="mt-0.5 flex flex-wrap gap-1">
          {serials.map((s, i) => (
            <span key={i} className="px-1.5 py-0.5 rounded bg-blue-50 text-blue-700 text-[10px] font-mono">{s}</span>
          ))}
        </div>
      )}
    </div>
  );
}
