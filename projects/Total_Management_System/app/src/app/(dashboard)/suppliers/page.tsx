'use client';

import { useState, useEffect } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useQueryClient } from '@tanstack/react-query';
import { useCustomers, useCreateCustomer } from '@/hooks/use-customers';
import { formatKRW, formatPhone } from '@/lib/utils/format';
import { EmptyState } from '@/components/ui/empty-state';
import { SearchInput } from '@/components/ui/search-input';
import { Building2, Users, GraduationCap, Plus, X, Pencil, Package, Save, Trash2, Search, Printer } from 'lucide-react';
import { CatalogPrintModal } from '@/components/purchasing/catalog-print-modal';
import { CollectPaymentModal } from '@/components/suppliers/collect-payment-modal';
import { useSupplierCatalog, useAddToCatalog, useUpdateCatalog, useRemoveFromCatalog } from '@/hooks/use-purchasing';
import { useProducts } from '@/hooks/use-sales';
import toast from 'react-hot-toast';
import type { Customer } from '@/lib/supabase/types';

const B2B_TABS = [
  { value: 'dealer', label: '딜러', icon: Users, color: 'purple', bgIcon: 'bg-purple-50', textIcon: 'text-purple-600' },
  { value: 'academy', label: '아카데미', icon: GraduationCap, color: 'emerald', bgIcon: 'bg-emerald-50', textIcon: 'text-emerald-600' },
  { value: 'supplier', label: '매입처', icon: Building2, color: 'amber', bgIcon: 'bg-amber-50', textIcon: 'text-amber-600' },
] as const;

const BADGE_STYLE: Record<string, string> = {
  dealer: 'bg-purple-100 text-purple-700',
  academy: 'bg-emerald-100 text-emerald-700',
  supplier: 'bg-amber-100 text-amber-700',
};

const BADGE_LABEL: Record<string, string> = {
  dealer: '딜러',
  academy: '아카데미',
  supplier: '매입처',
};

export default function B2BPartnersPage() {
  const [activeTab, setActiveTab] = useState<'dealer' | 'academy' | 'supplier'>('dealer');
  const [search, setSearch] = useState('');
  const [showAdd, setShowAdd] = useState(false);
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [isLg, setIsLg] = useState(false);

  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1024px)');
    setIsLg(mq.matches);
    const handler = (e: MediaQueryListEvent) => setIsLg(e.matches);
    mq.addEventListener('change', handler);
    return () => mq.removeEventListener('change', handler);
  }, []);

  const { data, isLoading } = useCustomers({ type: activeTab, search, limit: 100 });
  const partners = data?.customers || [];
  const selectedPartner = partners.find((p) => p.id === selectedId) || null;

  const tabConfig = B2B_TABS.find((t) => t.value === activeTab)!;

  const listContent = (
    <div className="space-y-3">
      {/* 서브탭 */}
      <div className="flex gap-1.5">
        {B2B_TABS.map((tab) => {
          const Icon = tab.icon;
          return (
            <button key={tab.value} onClick={() => { setActiveTab(tab.value); setSearch(''); setSelectedId(null); }}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold transition ${
                activeTab === tab.value ? 'bg-terracotta text-cream' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'
              }`}>
              <Icon size={12} />{tab.label}
            </button>
          );
        })}
      </div>

      <SearchInput value={search} onChange={setSearch} placeholder={`${tabConfig.label}명, 담당자명, 전화번호 검색`} />

      <Card padding={false}>
        {isLoading ? (
          <div className="p-4 space-y-3">{Array.from({ length: 4 }).map((_, i) => <Skeleton key={i} className="h-16 w-full" />)}</div>
        ) : partners.length === 0 ? (
          <EmptyState icon={tabConfig.icon} message={`등록된 ${tabConfig.label}가 없습니다`} />
        ) : (
          <div className="divide-y divide-neutral-100">
            {partners.map((p) => (
              <PartnerRow key={p.id} partner={p} tabConfig={tabConfig}
                selected={selectedId === p.id} onClick={() => setSelectedId(p.id)} />
            ))}
          </div>
        )}
      </Card>
      <p className="text-xs text-neutral-400">총 {partners.length}개 {tabConfig.label}</p>
    </div>
  );

  return (
    <>
      <Topbar title="B2B 거래처" action={
        <Button size="sm" onClick={() => setShowAdd(true)}><Plus size={14} />거래처 추가</Button>
      } />

      <div className="px-4 md:px-6 py-4">
        {isLg ? (
          <div className="flex gap-4">
            <div className="w-[480px] shrink-0">{listContent}</div>
            <div className="flex-1 min-w-0">
              {selectedPartner ? (
                <PartnerDetailPanel partner={selectedPartner} tabConfig={tabConfig} />
              ) : (
                <div className="flex items-center justify-center h-64 text-sm text-neutral-400">거래처를 선택해주세요</div>
              )}
            </div>
          </div>
        ) : (
          <>
            {listContent}
            <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="거래처 상세">
              {selectedPartner && <PartnerDetailPanel partner={selectedPartner} tabConfig={tabConfig} />}
            </SlidePanel>
          </>
        )}
      </div>

      {showAdd && <AddPartnerModal defaultType={activeTab} onClose={() => setShowAdd(false)} />}
    </>
  );
}

function PartnerRow({ partner: s, tabConfig, selected, onClick }: { partner: Customer; tabConfig: typeof B2B_TABS[number]; selected?: boolean; onClick?: () => void }) {
  return (
    <div onClick={onClick} className={`flex items-center gap-4 px-4 py-3 cursor-pointer transition ${selected ? 'bg-neutral-50 border-l-2 border-l-neutral-900' : 'hover:bg-warm-ivory/60'}`}>
      <div className={`w-9 h-9 rounded-lg ${tabConfig.bgIcon} flex items-center justify-center shrink-0`}>
        <tabConfig.icon size={18} className={tabConfig.textIcon} />
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex items-center gap-2">
          <span className="text-sm font-semibold text-indigo-black">{s.company_name || s.name}</span>
          <Badge className={BADGE_STYLE[s.customer_type] || BADGE_STYLE.supplier}>
            {BADGE_LABEL[s.customer_type] || s.customer_type}
          </Badge>
        </div>
        <div className="flex items-center gap-3 mt-0.5 text-xs text-neutral-500">
          {s.company_name && <span>담당: {s.name}</span>}
          {s.phone && <span>{s.phone}</span>}
          {s.memo && <span className="truncate max-w-[200px]">{s.memo}</span>}
        </div>
      </div>
      <div className="text-right shrink-0">
        {s.total_spent > 0 && (
          <p className="text-sm font-bold">{formatKRW(s.total_spent)}</p>
        )}
      </div>
    </div>
  );
}

/* ── 거래처 상세 패널 ── */
function PartnerDetailPanel({ partner: p, tabConfig }: { partner: Customer; tabConfig: typeof B2B_TABS[number] }) {
  const queryClient = useQueryClient();
  const [editing, setEditing] = useState(false);
  const [detailTab, setDetailTab] = useState<'info' | 'catalog'>('info');
  const [form, setForm] = useState({ name: '', company_name: '', phone: '', email: '', memo: '' });
  const [showCollect, setShowCollect] = useState(false);

  useEffect(() => {
    setForm({
      name: p.name || '', company_name: p.company_name || '',
      phone: p.phone || '', email: p.email || '',
      memo: p.memo || '',
    });
    setEditing(false);
    setDetailTab('info');
  }, [p.id]); // eslint-disable-line react-hooks/exhaustive-deps

  const handleSave = async () => {
    try {
      const res = await fetch(`/api/customers/${p.id}`, {
        method: 'PATCH', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(form),
      });
      if (!res.ok) throw new Error();
      toast.success('거래처 수정 완료');
      queryClient.invalidateQueries({ queryKey: ['customers'] });
      setEditing(false);
    } catch { toast.error('수정 실패'); }
  };

  return (
    <div className="bg-white rounded-xl border border-neutral-200 overflow-hidden">
      {/* 헤더 */}
      <div className="px-5 py-4 border-b border-neutral-100 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className={`w-10 h-10 rounded-lg ${tabConfig.bgIcon} flex items-center justify-center`}>
            <tabConfig.icon size={20} className={tabConfig.textIcon} />
          </div>
          <div>
            <h3 className="text-base font-bold text-indigo-black">{p.company_name || p.name}</h3>
            <Badge className={BADGE_STYLE[p.customer_type]}>{BADGE_LABEL[p.customer_type]}</Badge>
          </div>
        </div>
        <button onClick={() => setEditing(!editing)} className="text-sm text-neutral-500 hover:text-neutral-700">
          <Pencil size={14} />
        </button>
      </div>

      {/* 매입처일 때 탭 바 */}
      {p.customer_type === 'supplier' && (
        <div className="flex border-b border-neutral-100">
          <button
            onClick={() => setDetailTab('info')}
            className={`flex-1 py-2.5 text-xs font-semibold text-center transition ${detailTab === 'info' ? 'text-neutral-900 border-b-2 border-neutral-900' : 'text-neutral-400'}`}
          >기본정보</button>
          <button
            onClick={() => setDetailTab('catalog')}
            className={`flex-1 py-2.5 text-xs font-semibold text-center transition flex items-center justify-center gap-1 ${detailTab === 'catalog' ? 'text-neutral-900 border-b-2 border-neutral-900' : 'text-neutral-400'}`}
          ><Package size={12} />매입품목</button>
        </div>
      )}

      {/* 매입품목 탭 */}
      {detailTab === 'catalog' && p.customer_type === 'supplier' && (
        <SupplierCatalogSection supplierId={p.id} supplierName={p.company_name || p.name} />
      )}

      {/* 기본정보 탭 */}
      {detailTab === 'info' && <div className="px-5 py-4 space-y-4">
        {editing ? (
          <div className="space-y-3">
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-neutral-500">업체명</label>
                <input value={form.company_name} onChange={(e) => setForm({ ...form, company_name: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">담당자</label>
                <input value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">연락처</label>
                <input value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">이메일</label>
                <input value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })}
                  className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              </div>
            </div>
            <div>
              <label className="text-xs text-neutral-500">메모</label>
              <textarea value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })} rows={3}
                className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm" />
            </div>
            <div className="flex gap-2">
              <Button size="sm" onClick={handleSave}>저장</Button>
              <Button size="sm" variant="secondary" onClick={() => setEditing(false)}>취소</Button>
            </div>
          </div>
        ) : (
          <>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <div><p className="text-xs text-neutral-400">담당자</p><p className="font-medium">{p.name}</p></div>
              <div><p className="text-xs text-neutral-400">연락처</p><p className="font-medium">{p.phone ? formatPhone(p.phone) : '-'}</p></div>
              <div><p className="text-xs text-neutral-400">이메일</p><p>{p.email || '-'}</p></div>
              <div><p className="text-xs text-neutral-400">등록일</p><p>{p.created_at?.slice(0, 10) || '-'}</p></div>
            </div>
            {p.address_road && (
              <div className="text-sm">
                <p className="text-xs text-neutral-400">주소</p>
                <p>{p.postcode} {p.address_road} {p.address_detail}</p>
              </div>
            )}
            <div className="pt-3 border-t border-neutral-100">
              <p className="text-xs text-neutral-400 mb-1">메모</p>
              <div className="min-h-[40px] p-2 rounded-lg bg-neutral-50 text-sm text-neutral-600">
                {p.memo || <span className="text-neutral-300 italic">메모 없음</span>}
              </div>
            </div>
            {p.outstanding_balance > 0 && (
              <div className="pt-3 border-t border-neutral-100">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-xs text-neutral-400">미수금</p>
                    <p className="text-lg font-bold text-red-600">{formatKRW(p.outstanding_balance)}</p>
                  </div>
                  <button
                    onClick={() => setShowCollect(true)}
                    className="px-3 py-1.5 rounded-lg bg-neutral-900 text-white text-xs font-medium hover:bg-neutral-800 transition"
                  >
                    수금 처리
                  </button>
                </div>
              </div>
            )}
            {showCollect && (
              <CollectPaymentModal
                open={showCollect}
                customerId={p.id}
                customerName={p.company_name || p.name}
                onClose={() => setShowCollect(false)}
                onComplete={() => queryClient.invalidateQueries({ queryKey: ['customers'] })}
              />
            )}
          </>
        )}
      </div>}
    </div>
  );
}

/** 매입품목 카탈로그 섹션 */
function SupplierCatalogSection({ supplierId, supplierName }: { supplierId: string; supplierName: string }) {
  const { data, isLoading } = useSupplierCatalog(supplierId);
  const { data: products = [] } = useProducts();
  const addToCatalog = useAddToCatalog();
  const updateCatalog = useUpdateCatalog();
  const removeFromCatalog = useRemoveFromCatalog();
  const [showProductPicker, setShowProductPicker] = useState(false);
  const [productSearch, setProductSearch] = useState('');
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState({ order_name: '', features: '' });
  const [showCatalogPrint, setShowCatalogPrint] = useState(false);

  const catalog = data?.catalog || [];
  const catalogProductIds = new Set(catalog.map((c) => c.product_id));
  const availableProducts = products.filter((p) => !catalogProductIds.has(p.id) && p.category !== 'SUP');
  const filtered = productSearch.length >= 1
    ? availableProducts.filter((p) => p.name.toLowerCase().includes(productSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(productSearch.toLowerCase())).slice(0, 10)
    : availableProducts.slice(0, 10);

  function startEdit(entry: typeof catalog[0]) {
    setEditingId(entry.id);
    setEditForm({ order_name: entry.order_name, features: entry.features });
  }

  async function saveEdit() {
    if (!editingId) return;
    await updateCatalog.mutateAsync({ supplierId, catalogId: editingId, orderName: editForm.order_name, features: editForm.features });
    setEditingId(null);
  }

  async function handleAdd(productId: string) {
    await addToCatalog.mutateAsync({ supplierId, productIds: [productId] });
    setShowProductPicker(false);
    setProductSearch('');
  }

  async function handleRemove(catalogId: string) {
    await removeFromCatalog.mutateAsync({ supplierId, catalogId });
  }

  return (
    <div className="px-4 py-4 space-y-3">
      {/* 상단 버튼 */}
      <div className="flex items-center justify-between">
        <p className="text-xs font-semibold text-neutral-600">매입품목 ({catalog.length})</p>
        <div className="flex items-center gap-2">
          {catalog.length > 0 && (
            <button
              onClick={() => setShowCatalogPrint(true)}
              className="flex items-center gap-1 px-3 py-1.5 rounded-lg border border-neutral-200 text-neutral-600 text-xs font-medium hover:bg-neutral-50 transition"
            >
              <Printer size={12} />리스트 출력
            </button>
          )}
          <button
            onClick={() => setShowProductPicker(!showProductPicker)}
            className="flex items-center gap-1 px-3 py-1.5 rounded-lg bg-neutral-900 text-white text-xs font-medium hover:bg-neutral-800 transition"
          >
            <Plus size={12} />제품에서 불러오기
          </button>
        </div>
      </div>

      {/* 제품 선택 드롭다운 */}
      {showProductPicker && (
        <div className="rounded-lg border border-neutral-200 bg-white p-3 space-y-2">
          <div className="relative">
            <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={productSearch}
              onChange={(e) => setProductSearch(e.target.value)}
              placeholder="제품명 또는 SKU 검색"
              className="w-full h-8 pl-8 pr-3 rounded-lg border border-neutral-200 text-xs"
              autoFocus
            />
          </div>
          <div className="max-h-40 overflow-y-auto space-y-0.5">
            {filtered.length === 0 ? (
              <p className="text-xs text-neutral-400 text-center py-3">추가 가능한 제품 없음</p>
            ) : filtered.map((p) => (
              <button
                key={p.id}
                onClick={() => handleAdd(p.id)}
                className="w-full flex items-center justify-between px-3 py-2 rounded hover:bg-neutral-50 text-left"
              >
                <div>
                  <span className="text-xs font-medium">{p.name}</span>
                  {p.sku && !p.sku.startsWith('IW-') && <span className="text-[10px] text-neutral-400 ml-2">{p.sku}</span>}
                </div>
                {p.price_purchase > 0 && <span className="text-[10px] text-neutral-500">{formatKRW(p.price_purchase)}</span>}
              </button>
            ))}
          </div>
          <button onClick={() => { setShowProductPicker(false); setProductSearch(''); }} className="w-full py-1.5 text-xs text-neutral-400 hover:text-neutral-600">닫기</button>
        </div>
      )}

      {/* 카탈로그 목록 */}
      {isLoading ? (
        <div className="text-xs text-neutral-400 text-center py-8">로딩중...</div>
      ) : catalog.length === 0 ? (
        <div className="text-xs text-neutral-400 text-center py-8">등록된 매입품목이 없습니다<br />위 버튼으로 제품을 추가해주세요</div>
      ) : (
        <div className="space-y-2">
          {catalog.map((entry) => (
            <div key={entry.id} className="rounded-lg border border-neutral-200 p-3">
              {editingId === entry.id ? (
                /* 편집 모드 */
                <div className="space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold">{entry.product_name}</span>
                    <span className="text-[10px] text-neutral-400">{formatKRW(entry.price_purchase)}</span>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="text-[10px] text-neutral-400">주문명</label>
                      <input value={editForm.order_name} onChange={(e) => setEditForm({ ...editForm, order_name: e.target.value })}
                        placeholder="공장 발주용 품명"
                        className="w-full h-7 px-2 rounded border border-neutral-200 text-xs" />
                    </div>
                    <div>
                      <label className="text-[10px] text-neutral-400">특징</label>
                      <input value={editForm.features} onChange={(e) => setEditForm({ ...editForm, features: e.target.value })}
                        placeholder="규격, 특이사항 등"
                        className="w-full h-7 px-2 rounded border border-neutral-200 text-xs" />
                    </div>
                  </div>
                  <div className="flex gap-1.5 justify-end">
                    <button onClick={() => setEditingId(null)} className="px-2 py-1 text-[10px] text-neutral-500 hover:text-neutral-700">취소</button>
                    <button onClick={saveEdit} className="flex items-center gap-1 px-2 py-1 rounded bg-neutral-900 text-white text-[10px]">
                      <Save size={10} />저장
                    </button>
                  </div>
                </div>
              ) : (
                /* 읽기 모드 */
                <div className="flex items-start justify-between">
                  <div className="flex-1 min-w-0" onClick={() => startEdit(entry)} style={{ cursor: 'pointer' }}>
                    <div className="flex items-center gap-2 mb-1">
                      <span className="text-xs font-semibold truncate">{entry.product_name}</span>
                      <span className="text-[10px] text-neutral-400">{formatKRW(entry.price_purchase)}</span>
                    </div>
                    <div className="flex gap-4 text-[11px] text-neutral-500">
                      <span>주문명: {entry.order_name || <em className="text-neutral-300">미입력</em>}</span>
                      <span>특징: {entry.features || <em className="text-neutral-300">미입력</em>}</span>
                    </div>
                  </div>
                  <button onClick={() => handleRemove(entry.id)} className="text-neutral-300 hover:text-red-500 ml-2 shrink-0">
                    <Trash2 size={12} />
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>
      )}

      {/* 카탈로그 출력 모달 */}
      {showCatalogPrint && (
        <CatalogPrintModal supplierId={supplierId} supplierName={supplierName} onClose={() => setShowCatalogPrint(false)} />
      )}
    </div>
  );
}

function AddPartnerModal({ defaultType, onClose }: { defaultType: string; onClose: () => void }) {
  const createCustomer = useCreateCustomer();
  const [partnerType, setPartnerType] = useState(defaultType);
  const [form, setForm] = useState({
    companyName: '', name: '', phone: '', memo: '',
    businessNumber: '', representative: '', businessType: '', businessCategory: '',
    email: '', address: '', contactChannel: '',
  });

  const inputCls = "w-full h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40";

  async function handleSubmit() {
    if (!form.companyName.trim()) {
      toast.error('업체명을 입력해주세요');
      return;
    }
    await createCustomer.mutateAsync({
      name: form.name.trim() || form.companyName.trim(),
      company_name: form.companyName.trim(),
      phone: form.phone.trim() || undefined,
      email: form.email.trim() || undefined,
      address_road: form.address.trim() || undefined,
      memo: form.memo.trim() || undefined,
      customer_type: partnerType,
    });
    onClose();
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl w-full max-w-md mx-4 shadow-xl max-h-[90vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4 border-b border-neutral-100 sticky top-0 bg-white rounded-t-xl z-10">
          <h2 className="text-sm font-bold text-indigo-black">거래처 추가</h2>
          <button onClick={onClose} className="w-7 h-7 rounded-lg hover:bg-neutral-100 flex items-center justify-center">
            <X size={16} />
          </button>
        </div>
        <div className="px-5 py-4 space-y-3">
          {/* 거래처 유형 선택 */}
          <div>
            <label className="text-xs text-neutral-500">거래처 유형</label>
            <select value={partnerType} onChange={(e) => setPartnerType(e.target.value)} className={inputCls}>
              {B2B_TABS.map((t) => (
                <option key={t.value} value={t.value}>{t.label}</option>
              ))}
            </select>
          </div>

          <p className="text-[10px] text-neutral-400 font-semibold uppercase tracking-wider">기본 정보</p>
          <div className="grid grid-cols-2 gap-3">
            <div className="col-span-2">
              <label className="text-xs text-neutral-500">업체명 *</label>
              <input type="text" value={form.companyName} onChange={(e) => setForm({ ...form, companyName: e.target.value })}
                placeholder="업체명" autoFocus className={inputCls} />
            </div>
            {partnerType === 'supplier' && (
              <>
                <div>
                  <label className="text-xs text-neutral-500">사업자등록번호</label>
                  <input type="text" value={form.businessNumber} onChange={(e) => setForm({ ...form, businessNumber: e.target.value })}
                    placeholder="000-00-00000" className={inputCls} />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">대표자명</label>
                  <input type="text" value={form.representative} onChange={(e) => setForm({ ...form, representative: e.target.value })}
                    placeholder="대표자" className={inputCls} />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">업태</label>
                  <input type="text" value={form.businessType} onChange={(e) => setForm({ ...form, businessType: e.target.value })}
                    placeholder="제조, 도소매 등" className={inputCls} />
                </div>
                <div>
                  <label className="text-xs text-neutral-500">종목</label>
                  <input type="text" value={form.businessCategory} onChange={(e) => setForm({ ...form, businessCategory: e.target.value })}
                    placeholder="미용기기, 포장재 등" className={inputCls} />
                </div>
              </>
            )}
          </div>

          <p className="text-[10px] text-neutral-400 font-semibold uppercase tracking-wider pt-2">연락처</p>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500">담당자명</label>
              <input type="text" value={form.name} onChange={(e) => setForm({ ...form, name: e.target.value })}
                placeholder="담당자" className={inputCls} />
            </div>
            <div>
              <label className="text-xs text-neutral-500">전화번호</label>
              <input type="tel" value={form.phone} onChange={(e) => setForm({ ...form, phone: e.target.value })}
                placeholder="010-0000-0000" className={inputCls} />
            </div>
            <div>
              <label className="text-xs text-neutral-500">이메일</label>
              <input type="email" value={form.email} onChange={(e) => setForm({ ...form, email: e.target.value })}
                placeholder="email@example.com" className={inputCls} />
            </div>
            <div>
              <label className="text-xs text-neutral-500">연락 경로</label>
              <input type="text" value={form.contactChannel} onChange={(e) => setForm({ ...form, contactChannel: e.target.value })}
                placeholder="카톡, 전화 등" className={inputCls} />
            </div>
          </div>

          <div>
            <label className="text-xs text-neutral-500">사업장 주소</label>
            <input type="text" value={form.address} onChange={(e) => setForm({ ...form, address: e.target.value })}
              placeholder="사업장 주소" className={inputCls} />
          </div>

          <div>
            <label className="text-xs text-neutral-500">메모</label>
            <textarea value={form.memo} onChange={(e) => setForm({ ...form, memo: e.target.value })}
              placeholder="거래 조건, 결제 방식 등" rows={2}
              className="w-full px-3 py-2 rounded-lg border border-neutral-200 bg-warm-ivory text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 resize-none" />
          </div>
        </div>
        <div className="px-5 py-4 border-t border-neutral-100 flex gap-2 sticky bottom-0 bg-white rounded-b-xl">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={!form.companyName.trim() || createCustomer.isPending} onClick={handleSubmit}>
            {createCustomer.isPending ? '등록 중...' : '거래처 등록'}
          </Button>
        </div>
      </div>
    </div>
  );
}
