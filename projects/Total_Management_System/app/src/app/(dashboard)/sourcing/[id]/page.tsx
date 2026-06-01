'use client';

import { use, useMemo, useRef, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import {
  useSourcingDetail, useUpdateSourcingPo, useDeleteSourcingPo,
  useAddSourcingItem, useUpdateSourcingItem, useDeleteSourcingItem,
  useUploadItemImage, useDeleteItemImage, useLinkSourcingItem,
  type SourcingItem,
} from '@/hooks/use-sourcing';
import { RegisterLinkModal } from './_components/register-link-modals';
import { LabelPreview } from '@/app/(dashboard)/design-lab/_sections/sourcing-1688/LabelPreview';
import type { DemoPO } from '@/app/(dashboard)/design-lab/_sections/sourcing-1688/types';
import { printSourcingLabels, printSourcingPriceList } from '@/lib/sourcing/label-print';
import {
  ArrowLeft, Plus, Trash2, ExternalLink, Award, X, RotateCcw, Copy, PackageSearch,
  ImagePlus, Loader2, Printer, ListOrdered,
} from 'lucide-react';
import toast from 'react-hot-toast';

export default function SourcingDetailPage({ params }: { params: Promise<{ id: string }> }) {
  const { id } = use(params);
  const router = useRouter();
  const { data, isLoading } = useSourcingDetail(id);
  const updatePo = useUpdateSourcingPo();
  const deletePo = useDeleteSourcingPo();
  const addItem = useAddSourcingItem(id);
  const updateItem = useUpdateSourcingItem(id);
  const deleteItem = useDeleteSourcingItem(id);
  const uploadImg = useUploadItemImage(id);
  const delImg = useDeleteItemImage(id);
  const linkItem = useLinkSourcingItem(id);
  const [regModal, setRegModal] = useState<{ mode: 'product' | 'supply' | 'link'; item: SourcingItem } | null>(null);

  const po = data?.po;
  const items = useMemo(() => data?.items ?? [], [data]);

  // 라벨 컴포넌트(데모)용 DemoPO 매핑 — 수량은 라벨에서 미사용(0)
  const demoPo = useMemo<DemoPO | null>(() => {
    if (!po) return null;
    return {
      po_number: po.po_number,
      supplier_name: po.supplier_name ?? '',
      supplier_url: po.supplier_url ?? '',
      order_date: po.order_date,
      exchange_rate: po.exchange_rate,
      items: items.map((it) => ({
        id: it.id,
        vendor_url: it.vendor_url ?? '',
        product_name: it.product_name,
        features_memo: it.features_memo ?? '',
        moq: it.moq,
        unit_price: it.unit_price,
        quantity: 0,
        sticker_no: it.sticker_no,
        inbound_photos: it.inbound_photos ?? [],
        inbound_memo: it.inbound_memo ?? '',
        inspection_status: it.inspection_status === 'selected' ? 'promoted' : it.inspection_status,
      })),
    };
  }, [po, items]);

  const selected = items.filter((it) => it.inspection_status === 'selected');

  // 업체별 채택 현황 — "어느 공장이 좋은 제품을 많이 내놓나"
  const vendorSummary = useMemo(() => {
    const map = new Map<string, { total: number; selected: number; rejected: number }>();
    for (const it of items) {
      const key = (it.supplier_name || '').trim() || '(회사명 미입력)';
      const cur = map.get(key) ?? { total: 0, selected: 0, rejected: 0 };
      cur.total++;
      if (it.inspection_status === 'selected') cur.selected++;
      if (it.inspection_status === 'rejected') cur.rejected++;
      map.set(key, cur);
    }
    return Array.from(map.entries())
      .map(([name, c]) => ({ name, ...c }))
      .sort((a, b) => b.selected - a.selected || b.total - a.total);
  }, [items]);

  if (isLoading || !po) {
    return (
      <div className="min-h-screen bg-warm-ivory">
        <Topbar title="샘플 소싱" />
        <div className="p-6 text-center text-sm text-neutral-400">불러오는 중…</div>
      </div>
    );
  }

  const handleDelete = async () => {
    if (!confirm(`${po.po_number} 소싱을 삭제할까요? (품목 전부 삭제)`)) return;
    await deletePo.mutateAsync(id);
    router.push('/sourcing');
  };

  // A4 가격 리스트 (번호·품목명·업체·¥단가·₩가격) — 사장님 참고용 전체 출력
  const printPriceList = () => {
    if (items.length === 0) return;
    printSourcingPriceList(
      items.map((it) => ({
        seq: it.sticker_no.split('-').pop() || '',
        product_name: it.product_name,
        unit_price: it.unit_price,
        supplier_name: it.supplier_name,
      })),
      po.exchange_rate,
      po.memo || po.po_number
    );
  };

  const copySelectionList = () => {
    if (selected.length === 0) return;
    const lines = selected.map((it) => {
      const cols = [
        it.supplier_name || '',
        it.supplier_url || '',
        it.product_name,
        it.vendor_url || '',
        `¥${it.unit_price}`,
        it.features_memo || '',
      ];
      return cols.join('\t');
    });
    const text = ['회사명\t회사링크\t품목명\t품목링크\t단가(CNY)\t특징', ...lines].join('\n');
    navigator.clipboard.writeText(text).then(
      () => toast.success(`선별 ${selected.length}건 복사됨 (탭 구분 — 엑셀/시트에 붙여넣기)`),
      () => toast.error('복사 실패')
    );
  };

  return (
    <div className="min-h-screen bg-warm-ivory pb-16">
      <Topbar title="샘플 소싱" />
      <div className="p-4 sm:p-6 max-w-5xl mx-auto space-y-5">
        {/* 헤더 — 소싱 회차 (매입처는 제품별 1688 링크로 분산, 단일 매입처 헤더 제거) */}
        <div className="flex items-center gap-3">
          <button type="button" onClick={() => router.push('/sourcing')} className="text-neutral-500 hover:text-indigo-black">
            <ArrowLeft size={20} />
          </button>
          <div className="flex-1 min-w-0">
            <div className="font-mono text-base font-bold text-indigo-black">{po.po_number}</div>
            <div className="text-[11px] text-neutral-400">{po.order_date} · 소싱 회차</div>
          </div>
          <Button variant="secondary" size="sm" onClick={printPriceList}>
            <ListOrdered size={14} className="mr-1" /> 가격 리스트 (A4)
          </Button>
          <Button variant="ghost" size="sm" onClick={handleDelete} className="text-rose-500">
            <Trash2 size={14} className="mr-1" /> 소싱 삭제
          </Button>
        </div>

        {/* 회차명 + 환율 (환율 → 라벨 한화 가격에 즉시 반영) */}
        <div className="flex items-center gap-2 flex-wrap">
          <input
            type="text"
            defaultValue={po.memo ?? ''}
            placeholder="소싱 회차명 (선택) — 예: 5월 가위 1차"
            onBlur={(e) => e.target.value !== (po.memo ?? '') && updatePo.mutate({ id, memo: e.target.value })}
            className="flex-1 min-w-[180px] px-3 py-2 text-sm rounded-lg border border-neutral-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-black/10 focus:border-neutral-400"
          />
          <div className="flex items-center gap-1.5 px-3 py-2 rounded-lg border border-neutral-200 bg-white" title="라벨 한화 가격에 적용되는 환율">
            <span className="text-[11px] text-neutral-500 whitespace-nowrap">환율 1¥ =</span>
            <input
              type="number"
              defaultValue={po.exchange_rate}
              onBlur={(e) => { const v = Number(e.target.value) || 0; if (v !== po.exchange_rate) updatePo.mutate({ id, exchange_rate: v }); }}
              className="w-16 text-sm text-right font-bold text-indigo-black focus:outline-none"
            />
            <span className="text-[11px] text-neutral-500">원</span>
          </div>
        </div>

        {/* STEP 1. 품목 */}
        <Section n={1} title="품목" sub="회사·제품이미지(1688 붙여넣기)·품목링크·단가·특징 입력 · 복제로 같은 업체 빠르게">
          <div className="space-y-3">
            {items.map((it, idx) => (
              <ItemRow
                key={it.id}
                item={it}
                idx={idx}
                exchangeRate={po.exchange_rate}
                onAddImage={(file) => uploadImg.mutateAsync({ itemId: it.id, file })}
                onRemoveImage={(url) => delImg.mutate({ itemId: it.id, url })}
                onPatch={(patch) => updateItem.mutate({ itemId: it.id, ...patch })}
                onDuplicate={() => addItem.mutate({
                  supplier_name: it.supplier_name,
                  supplier_url: it.supplier_url,
                  vendor_url: it.vendor_url,
                  product_name: it.product_name,
                  features_memo: it.features_memo,
                  unit_price: it.unit_price,
                  moq: it.moq,
                })}
                onDelete={() => deleteItem.mutate(it.id)}
              />
            ))}
            <button
              type="button"
              onClick={() => addItem.mutate(undefined)}
              className="w-full py-2.5 rounded-lg border border-dashed border-neutral-300 text-sm text-neutral-500 hover:bg-neutral-50 inline-flex items-center justify-center gap-1.5"
            >
              <Plus size={14} /> 품목 추가
            </button>
          </div>
        </Section>

        {/* STEP 2. 라벨 */}
        <Section n={2} title="라벨 인쇄" sub="QR + 번호 + 품목명 · 라벨프린터로 출력해 샘플에 부착">
          {demoPo && (
            <LabelPreview
              po={demoPo}
              qrBaseUrl={`${typeof window !== 'undefined' ? window.location.origin : 'https://app-eta-sandy-75.vercel.app'}/sourcing/inbound`}
            />
          )}
        </Section>

        {/* STEP 3. 선별 */}
        <Section n={3} title="선별 (실테스트 후)" sub="제품 이미지 보고 채택 / 탈락 · 업체별 채택 현황으로 좋은 공장 선정 (도착품은 라벨로 매칭)">
          {items.length === 0 ? (
            <Empty>STEP 1에서 품목을 먼저 입력하세요.</Empty>
          ) : (
            <>
              {/* 업체별 채택 현황 */}
              <div className="mb-3 flex flex-wrap gap-2">
                {vendorSummary.map((v) => (
                  <div key={v.name} className="inline-flex items-center gap-2 px-2.5 py-1.5 rounded-lg bg-neutral-100 border border-neutral-200 text-xs">
                    <span className="font-bold text-indigo-black truncate max-w-[160px]">{v.name}</span>
                    <span className="text-emerald-700 font-bold">채택 {v.selected}</span>
                    <span className="text-neutral-400">/ {v.total}</span>
                    {v.rejected > 0 && <span className="text-rose-400">탈락 {v.rejected}</span>}
                  </div>
                ))}
              </div>
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
                {items.map((it) => (
                  <SelectionCell
                    key={it.id}
                    item={it}
                    onSelect={() => updateItem.mutate({ itemId: it.id, inspection_status: 'selected' })}
                    onReject={() => updateItem.mutate({ itemId: it.id, inspection_status: 'rejected' })}
                    onReset={() => updateItem.mutate({ itemId: it.id, inspection_status: 'pending' })}
                  />
                ))}
              </div>
            </>
          )}
        </Section>

        {/* STEP 4. 선별 리스트 */}
        <Section n={4} title={`선별 리스트 (채택 ${selected.length}건)`} sub="채택분 — 사장님이 이 리스트 보고 아임웹/TMS에 직접 등록">
          {selected.length === 0 ? (
            <Empty>아직 채택된 제품이 없습니다. STEP 3에서 [채택]을 누르세요.</Empty>
          ) : (
            <div className="space-y-2">
              <div className="flex justify-end">
                <Button size="sm" variant="secondary" onClick={copySelectionList}>
                  <Copy size={13} className="mr-1" /> 리스트 복사 (엑셀/시트)
                </Button>
              </div>
              {selected.map((it) => (
                <div key={it.id} className="flex items-start gap-3 p-3 rounded-lg bg-white border border-neutral-200">
                  {it.inbound_photos?.[0] ? (
                    <div className="w-12 h-12 rounded-md flex-shrink-0 bg-cover bg-center border border-neutral-200"
                      style={{ background: it.inbound_photos[0].startsWith('mock:') ? it.inbound_photos[0].split(':').slice(2).join(':') : `url(${it.inbound_photos[0]}) center/cover` }} />
                  ) : (
                    <div className="w-12 h-12 rounded-md flex-shrink-0 bg-neutral-100 flex items-center justify-center text-neutral-300">
                      <PackageSearch size={18} />
                    </div>
                  )}
                  <div className="flex-1 min-w-0">
                    {it.supplier_name && (
                      <div className="text-[11px] text-neutral-500 mb-0.5 inline-flex items-center gap-1 max-w-full">
                        <span className="truncate">🏢 {it.supplier_name}</span>
                        {it.supplier_url && (
                          <a href={it.supplier_url} target="_blank" rel="noreferrer" className="text-blue-600 flex-shrink-0"><ExternalLink size={9} /></a>
                        )}
                      </div>
                    )}
                    <div className="text-sm font-bold text-indigo-black truncate">{it.product_name || '(품목명 없음)'}</div>
                    <div className="text-[11px] text-neutral-500 mt-0.5">¥{it.unit_price}{it.features_memo ? ` · ${it.features_memo}` : ''}</div>
                    {it.vendor_url && (
                      <a href={it.vendor_url} target="_blank" rel="noreferrer" className="text-[11px] text-blue-600 inline-flex items-center gap-0.5 mt-0.5">
                        품목 링크 <ExternalLink size={10} />
                      </a>
                    )}
                  </div>
                  <div className="flex flex-col items-end gap-1.5 flex-shrink-0">
                    <span className="font-mono text-[10px] text-neutral-400">{it.sticker_no.split('-').pop()}</span>
                    {it.linked_product_id ? (
                      <>
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700 text-[10px] font-bold border border-emerald-200">
                          ✅ {it.linked_product?.sku || (it.linked_product?.category === 'SUP' ? '부자재' : '등록됨')}
                        </span>
                        <div className="flex items-center gap-1.5">
                          <button type="button" onClick={() => router.push(it.linked_product?.category === 'SUP' ? '/supplies' : '/products')} className="text-[10px] text-blue-600 hover:underline">보기</button>
                          <button type="button" onClick={() => { if (confirm('연결을 해제할까요? (등록된 제품은 그대로 남습니다)')) linkItem.mutate({ itemId: it.id, linked_product_id: null }); }} className="text-[10px] text-neutral-400 hover:text-rose-500">연결해제</button>
                        </div>
                      </>
                    ) : (
                      <div className="flex flex-col items-end gap-1">
                        <button type="button" onClick={() => setRegModal({ mode: 'product', item: it })} className="text-[10px] px-2 py-1 rounded-md bg-indigo-black text-white font-bold hover:opacity-90">제품등록</button>
                        <button type="button" onClick={() => setRegModal({ mode: 'supply', item: it })} className="text-[10px] px-2 py-1 rounded-md bg-neutral-100 text-neutral-700 border border-neutral-200 hover:bg-neutral-200">부자재등록</button>
                        <button type="button" onClick={() => setRegModal({ mode: 'link', item: it })} className="text-[10px] text-blue-600 hover:underline">기존 연결</button>
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          )}
        </Section>
      </div>

      {regModal && (
        <RegisterLinkModal
          mode={regModal.mode}
          item={regModal.item}
          poId={id}
          exchangeRate={po.exchange_rate}
          onClose={() => setRegModal(null)}
        />
      )}
    </div>
  );
}

// ── 하위 컴포넌트 ─────────────────────────────────────────

function Section({ n, title, sub, children }: { n: number; title: string; sub: string; children: React.ReactNode }) {
  return (
    <section className="rounded-xl border border-neutral-200 bg-white overflow-hidden">
      <div className="px-4 py-3 border-b border-neutral-100 flex items-center gap-3">
        <span className="inline-flex items-center justify-center w-8 h-8 rounded-full bg-indigo-black text-white text-sm font-bold flex-shrink-0">{n}</span>
        <div className="min-w-0">
          <div className="font-bold text-indigo-black text-sm">{title}</div>
          <div className="text-[11px] text-neutral-500 truncate">{sub}</div>
        </div>
      </div>
      <div className="p-4">{children}</div>
    </section>
  );
}

function Empty({ children }: { children: React.ReactNode }) {
  return <div className="text-center py-8 text-xs text-neutral-400">{children}</div>;
}

const PRINT_SIZES = [
  { id: 'p30x15', w: 30, h: 15, label: '30 × 15' },
  { id: 'p40x20', w: 40, h: 20, label: '40 × 20' },
  { id: 'p40x30', w: 40, h: 30, label: '40 × 30' },
];

/** 품목별 즉시 라벨 인쇄 칩 (수량+사이즈 팝오버 → 그 품목 라벨만 인쇄) */
function LabelPrintChip({ item }: { item: SourcingItem }) {
  const [open, setOpen] = useState(false);
  const [qty, setQty] = useState(1);
  const [sizeId, setSizeId] = useState('p40x20');

  const doPrint = () => {
    const s = PRINT_SIZES.find((x) => x.id === sizeId) || PRINT_SIZES[1];
    const base = `${typeof window !== 'undefined' ? window.location.origin : 'https://app-eta-sandy-75.vercel.app'}/sourcing/inbound`;
    printSourcingLabels(
      [{
        sticker_no: item.sticker_no,
        product_name: item.product_name,
        qrValue: `${base}/${item.id}`,
        copies: Math.max(1, qty),
      }],
      { w: s.w, h: s.h }
    );
    setOpen(false);
  };

  return (
    <div className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        title="이 품목 라벨 인쇄"
        className="inline-flex items-center gap-1 px-2 py-1 rounded-md bg-indigo-black/5 text-indigo-black text-[10px] font-bold border border-neutral-200 hover:bg-indigo-black/10"
      >
        <Printer size={12} /> 라벨인쇄
      </button>
      {open && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div className="absolute right-0 top-full mt-1 z-20 w-44 rounded-lg border border-neutral-200 bg-white shadow-lg p-2.5 space-y-2">
            <div className="text-[10px] text-neutral-500 font-mono truncate">{item.sticker_no}</div>
            <label className="block">
              <span className="text-[10px] text-neutral-500 block mb-0.5">사이즈 (mm)</span>
              <select value={sizeId} onChange={(e) => setSizeId(e.target.value)} className="w-full px-2 py-1.5 rounded-lg border border-neutral-200 bg-white text-xs focus:outline-none focus:ring-2 focus:ring-indigo-black/10">
                {PRINT_SIZES.map((s) => <option key={s.id} value={s.id}>{s.label}</option>)}
              </select>
            </label>
            <label className="block">
              <span className="text-[10px] text-neutral-500 block mb-0.5">수량 (매수)</span>
              <input type="number" min={1} value={qty} onChange={(e) => setQty(Math.max(1, Number(e.target.value) || 1))} className="w-full px-2 py-1.5 rounded-lg border border-neutral-200 bg-white text-sm font-bold focus:outline-none focus:ring-2 focus:ring-indigo-black/10" />
            </label>
            <div className="flex gap-1.5 pt-0.5">
              <button type="button" onClick={() => setOpen(false)} className="flex-1 text-[11px] py-1.5 rounded-lg bg-neutral-100 text-neutral-600 hover:bg-neutral-200">취소</button>
              <button type="button" onClick={doPrint} className="flex-1 text-[11px] py-1.5 rounded-lg bg-indigo-black text-white font-bold hover:opacity-90 inline-flex items-center justify-center gap-1"><Printer size={11} /> 인쇄</button>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function ItemRow({ item, idx, exchangeRate, onAddImage, onRemoveImage, onPatch, onDuplicate, onDelete }: {
  item: SourcingItem;
  idx: number;
  exchangeRate: number;
  onAddImage: (file: File) => Promise<unknown>;
  onRemoveImage: (url: string) => void;
  onPatch: (patch: Partial<SourcingItem>) => void;
  onDuplicate: () => void;
  onDelete: () => void;
}) {
  const cls = 'w-full px-2.5 py-1.5 text-sm rounded-lg border border-neutral-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-black/10 focus:border-neutral-400';
  const krw = Math.round((item.unit_price || 0) * (exchangeRate || 0));
  return (
    <div className="rounded-lg border border-neutral-200 bg-neutral-50/50 p-3">
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-2">
          <span className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-indigo-black text-white text-[11px] font-bold">{String(idx + 1).padStart(2, '0')}</span>
          <span className="font-mono text-[11px] text-neutral-500">{item.sticker_no}</span>
        </div>
        <div className="flex items-center gap-1.5">
          <LabelPrintChip item={item} />
          <button type="button" onClick={onDuplicate} title="복제 (회사 정보 유지 · 새 바코드)" className="text-neutral-400 hover:text-indigo-black p-0.5"><Copy size={14} /></button>
          <button type="button" onClick={onDelete} title="삭제" className="text-neutral-400 hover:text-rose-500 p-0.5"><Trash2 size={14} /></button>
        </div>
      </div>

      {/* 업체 (회사명 · 회사링크) — 복제 시 유지되는 부분 */}
      <div className="grid grid-cols-1 md:grid-cols-12 gap-2 mb-2 pb-2.5 border-b border-dashed border-neutral-300">
        <div className="md:col-span-5">
          <input type="text" defaultValue={item.supplier_name ?? ''} placeholder="회사명 (예: 光达美容工具)"
            onBlur={(e) => e.target.value !== (item.supplier_name ?? '') && onPatch({ supplier_name: e.target.value })} className={cls} />
        </div>
        <div className="md:col-span-7 relative">
          <input type="url" defaultValue={item.supplier_url ?? ''} placeholder="회사 링크 (https://shop....1688.com)"
            onBlur={(e) => e.target.value !== (item.supplier_url ?? '') && onPatch({ supplier_url: e.target.value })} className={cls + ' pr-8'} />
          {item.supplier_url && (
            <a href={item.supplier_url} target="_blank" rel="noreferrer" className="absolute right-2 top-1/2 -translate-y-1/2 text-neutral-400 hover:text-indigo-black"><ExternalLink size={13} /></a>
          )}
        </div>
      </div>

      {/* 제품 이미지 — 주문 시 1688 이미지 붙여넣기/업로드 (서버 영구 저장) */}
      <ImageZone images={item.inbound_photos ?? []} onAdd={onAddImage} onRemove={onRemoveImage} />

      {/* 품목 (품목링크 · 품목명 · 단가 · MOQ · 특징) */}
      <div className="grid grid-cols-1 md:grid-cols-12 gap-2">
        <div className="md:col-span-12 relative">
          <input type="url" defaultValue={item.vendor_url ?? ''} placeholder="품목 링크 (https://detail.1688.com/offer/...)"
            onBlur={(e) => e.target.value !== (item.vendor_url ?? '') && onPatch({ vendor_url: e.target.value })} className={cls + ' pr-8'} />
          {item.vendor_url && (
            <a href={item.vendor_url} target="_blank" rel="noreferrer" className="absolute right-2 top-1/2 -translate-y-1/2 text-neutral-400 hover:text-indigo-black"><ExternalLink size={13} /></a>
          )}
        </div>
        <div className="md:col-span-6">
          <input type="text" defaultValue={item.product_name} placeholder="품목명 (예: 6.0인치 일자 가위)"
            onBlur={(e) => e.target.value !== item.product_name && onPatch({ product_name: e.target.value })} className={cls} />
        </div>
        <div className="md:col-span-3">
          <input type="number" step="0.01" defaultValue={item.unit_price || ''} placeholder="단가 ¥"
            onBlur={(e) => { const v = Number(e.target.value) || 0; if (v !== item.unit_price) onPatch({ unit_price: v }); }} className={cls} />
          {item.unit_price > 0 && (
            <div className="mt-1 text-[11px] text-emerald-700 font-medium text-right">≈ ₩{krw.toLocaleString()}</div>
          )}
        </div>
        <div className="md:col-span-3">
          <input type="number" defaultValue={item.moq ?? ''} placeholder="MOQ"
            onBlur={(e) => { const v = e.target.value ? Number(e.target.value) : null; if (v !== item.moq) onPatch({ moq: v }); }} className={cls} />
        </div>
        <div className="md:col-span-12">
          <input type="text" defaultValue={item.features_memo ?? ''} placeholder="특징 메모 (예: 날 광택 양호, 풀너트)"
            onBlur={(e) => e.target.value !== (item.features_memo ?? '') && onPatch({ features_memo: e.target.value })} className={cls} />
        </div>
      </div>
    </div>
  );
}

function ImageZone({ images, onAdd, onRemove }: {
  images: string[];
  onAdd: (file: File) => Promise<unknown>;
  onRemove: (url: string) => void;
}) {
  const fileRef = useRef<HTMLInputElement>(null);
  const [busy, setBusy] = useState(false);

  const upload = async (files: File[]) => {
    const imgs = files.filter((f) => f.type.startsWith('image/'));
    if (imgs.length === 0) return;
    setBusy(true);
    try {
      for (const f of imgs) {
        if (images.length >= 5) break;
        await onAdd(f);
      }
    } finally {
      setBusy(false);
    }
  };

  const onPaste = (e: React.ClipboardEvent) => {
    const imgs = Array.from(e.clipboardData.items)
      .filter((i) => i.type.startsWith('image/'))
      .map((i) => i.getAsFile())
      .filter((f): f is File => !!f);
    if (imgs.length) {
      e.preventDefault();
      upload(imgs);
    }
  };

  return (
    <div className="mb-2">
      <div
        tabIndex={0}
        onPaste={onPaste}
        className="rounded-lg border border-dashed border-neutral-300 p-2 outline-none focus:border-indigo-black focus:ring-2 focus:ring-indigo-black/10"
      >
        <div className="flex flex-wrap items-center gap-2">
          {images.map((url, i) => (
            <div key={i} className="relative w-16 h-16 rounded-md overflow-hidden border border-neutral-200 bg-cover bg-center" style={{ backgroundImage: `url(${url})` }}>
              <button type="button" onClick={() => onRemove(url)} className="absolute top-0.5 right-0.5 w-4 h-4 bg-black/60 text-white rounded-full flex items-center justify-center text-[10px] leading-none">×</button>
            </div>
          ))}
          {images.length < 5 && (
            <button
              type="button"
              onClick={() => fileRef.current?.click()}
              disabled={busy}
              className="w-16 h-16 rounded-md border border-dashed border-neutral-300 flex flex-col items-center justify-center text-neutral-400 hover:bg-neutral-50 disabled:opacity-50"
            >
              {busy ? <Loader2 size={16} className="animate-spin" /> : <><ImagePlus size={16} /><span className="text-[8px] mt-0.5">추가</span></>}
            </button>
          )}
          <span className="text-[10px] text-neutral-400 ml-1 leading-tight">
            제품 이미지 — 여기 <strong>클릭 후 Ctrl+V</strong> (1688 이미지 붙여넣기) 또는 [+] 업로드
          </span>
        </div>
      </div>
      <input
        ref={fileRef}
        type="file"
        accept="image/*"
        multiple
        className="hidden"
        onChange={(e) => { upload(Array.from(e.target.files ?? [])); if (fileRef.current) fileRef.current.value = ''; }}
      />
    </div>
  );
}

function SelectionCell({ item, onSelect, onReject, onReset }: {
  item: SourcingItem; onSelect: () => void; onReject: () => void; onReset: () => void;
}) {
  const seq = item.sticker_no.split('-').pop();
  const photo = item.inbound_photos?.[0];
  const st = item.inspection_status;
  const border =
    st === 'selected' ? 'border-emerald-300 bg-emerald-50/30'
      : st === 'rejected' ? 'border-rose-200 bg-rose-50/30 opacity-60'
        : 'border-neutral-200';
  return (
    <div className={`group relative rounded-xl border-2 overflow-hidden ${border}`}>
      <div className="aspect-square relative">
        {photo ? (
          <div className="absolute inset-0" style={{ background: photo.startsWith('mock:') ? photo.split(':').slice(2).join(':') : `url(${photo}) center/cover` }} />
        ) : (
          <div className="absolute inset-0 flex items-center justify-center text-neutral-300"><PackageSearch size={28} /></div>
        )}
        <span className="absolute top-2 left-2 inline-flex items-center justify-center min-w-[26px] h-6 px-1.5 rounded-full bg-white/95 text-indigo-black text-[11px] font-bold shadow-sm">{seq}</span>
        <div className="absolute bottom-0 left-0 right-0 p-2 bg-gradient-to-t from-black/75 to-transparent">
          <div className="text-[11px] font-bold text-white truncate">{item.product_name || '(품목명 없음)'}</div>
          {item.supplier_name && <div className="text-[9px] text-white/70 truncate">🏢 {item.supplier_name}</div>}
        </div>
      </div>
      <div className="p-1.5 flex items-center gap-1 justify-center">
        {st !== 'selected' && st !== 'rejected' && (
          <>
            <button type="button" onClick={onSelect} className="flex-1 inline-flex items-center justify-center gap-1 px-2 py-1 rounded-md bg-emerald-600 text-white text-[11px] font-bold hover:bg-emerald-700"><Award size={11} /> 채택</button>
            <button type="button" onClick={onReject} className="inline-flex items-center justify-center gap-1 px-2 py-1 rounded-md bg-white text-neutral-600 text-[11px] border border-neutral-200 hover:bg-neutral-50"><X size={11} /> 탈락</button>
          </>
        )}
        {(st === 'selected' || st === 'rejected') && (
          <button type="button" onClick={onReset} className="flex-1 inline-flex items-center justify-center gap-1 px-2 py-1 rounded-md bg-white text-neutral-600 text-[11px] border border-neutral-200 hover:bg-neutral-50"><RotateCcw size={11} /> 되돌리기</button>
        )}
      </div>
    </div>
  );
}
