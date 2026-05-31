'use client';

import { use, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Button } from '@/components/ui/button';
import {
  useSourcingDetail, useUpdateSourcingPo, useDeleteSourcingPo,
  useAddSourcingItem, useUpdateSourcingItem, useDeleteSourcingItem,
  type SourcingItem,
} from '@/hooks/use-sourcing';
import { LabelPreview } from '@/app/(dashboard)/design-lab/_sections/sourcing-1688/LabelPreview';
import type { DemoPO } from '@/app/(dashboard)/design-lab/_sections/sourcing-1688/types';
import {
  ArrowLeft, Plus, Trash2, ExternalLink, Award, X, RotateCcw, Copy, Printer, PackageSearch,
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

  if (isLoading || !po) {
    return (
      <div className="min-h-screen bg-warm-ivory">
        <Topbar title="샘플 소싱" />
        <div className="p-6 text-center text-sm text-neutral-400">불러오는 중…</div>
      </div>
    );
  }

  const handleDelete = async () => {
    if (!confirm(`${po.po_number} 발주를 삭제할까요? (품목 전부 삭제)`)) return;
    await deletePo.mutateAsync(id);
    router.push('/sourcing');
  };

  const copySelectionList = () => {
    if (selected.length === 0) return;
    const lines = selected.map((it) => {
      const cols = [
        it.product_name,
        it.vendor_url || '',
        `¥${it.unit_price}`,
        it.features_memo || '',
      ];
      return cols.join('\t');
    });
    const text = ['품목명\t1688링크\t단가(CNY)\t특징', ...lines].join('\n');
    navigator.clipboard.writeText(text).then(
      () => toast.success(`선별 ${selected.length}건 복사됨 (탭 구분 — 엑셀/시트에 붙여넣기)`),
      () => toast.error('복사 실패')
    );
  };

  return (
    <div className="min-h-screen bg-warm-ivory pb-16">
      <Topbar title="샘플 소싱" />
      <div className="p-4 sm:p-6 max-w-5xl mx-auto space-y-5">
        {/* 헤더 */}
        <div className="flex items-center gap-3">
          <button type="button" onClick={() => router.push('/sourcing')} className="text-neutral-500 hover:text-indigo-black">
            <ArrowLeft size={20} />
          </button>
          <div className="flex-1">
            <div className="font-mono text-base font-bold text-indigo-black">{po.po_number}</div>
            <div className="text-[11px] text-neutral-400">{po.order_date}</div>
          </div>
          <Button variant="ghost" size="sm" onClick={handleDelete} className="text-rose-500">
            <Trash2 size={14} className="mr-1" /> 발주 삭제
          </Button>
        </div>

        {/* 매입처 */}
        <section className="rounded-xl border border-neutral-200 bg-white p-4">
          <div className="text-[11px] font-bold text-neutral-500 uppercase tracking-wider mb-3">매입처 (1688)</div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <LabeledInput label="회사명 (중문 가능)" defaultValue={po.supplier_name ?? ''} placeholder="예) 光达美容工具"
              onBlur={(v) => v !== (po.supplier_name ?? '') && updatePo.mutate({ id, supplier_name: v })} />
            <LabeledInput label="회사 홈 URL" defaultValue={po.supplier_url ?? ''} placeholder="https://shop....1688.com"
              onBlur={(v) => v !== (po.supplier_url ?? '') && updatePo.mutate({ id, supplier_url: v })} />
          </div>
        </section>

        {/* STEP 1. 품목 */}
        <Section n={1} title="품목" sub="1688 링크·품목명·단가·특징 입력 (수량 입력 없음 — 선별용)">
          <div className="space-y-3">
            {items.map((it, idx) => (
              <ItemRow
                key={it.id}
                item={it}
                idx={idx}
                onPatch={(patch) => updateItem.mutate({ itemId: it.id, ...patch })}
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
          {demoPo && <LabelPreview po={demoPo} />}
        </Section>

        {/* STEP 3. 선별 */}
        <Section n={3} title="선별 (실테스트 후)" sub="실테스트 결과로 채택 / 탈락 결정">
          {items.length === 0 ? (
            <Empty>STEP 1에서 품목을 먼저 입력하세요.</Empty>
          ) : (
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
                    <div className="text-sm font-bold text-indigo-black truncate">{it.product_name || '(품목명 없음)'}</div>
                    <div className="text-[11px] text-neutral-500 mt-0.5">¥{it.unit_price}{it.features_memo ? ` · ${it.features_memo}` : ''}</div>
                    {it.vendor_url && (
                      <a href={it.vendor_url} target="_blank" rel="noreferrer" className="text-[11px] text-blue-600 inline-flex items-center gap-0.5 mt-0.5">
                        1688 상품 링크 <ExternalLink size={10} />
                      </a>
                    )}
                  </div>
                  <span className="font-mono text-[10px] text-neutral-400">{it.sticker_no.split('-').pop()}</span>
                </div>
              ))}
            </div>
          )}
        </Section>
      </div>
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

function LabeledInput({ label, defaultValue, placeholder, onBlur }: {
  label: string; defaultValue: string; placeholder?: string; onBlur: (v: string) => void;
}) {
  return (
    <label className="block">
      <span className="block text-[11px] text-neutral-500 mb-1 font-medium">{label}</span>
      <input
        type="text"
        defaultValue={defaultValue}
        placeholder={placeholder}
        onBlur={(e) => onBlur(e.target.value)}
        className="w-full px-3 py-2 text-sm rounded-lg border border-neutral-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-black/10 focus:border-neutral-400"
      />
    </label>
  );
}

function ItemRow({ item, idx, onPatch, onDelete }: {
  item: SourcingItem;
  idx: number;
  onPatch: (patch: Partial<SourcingItem>) => void;
  onDelete: () => void;
}) {
  const cls = 'w-full px-2.5 py-1.5 text-sm rounded-lg border border-neutral-200 bg-white focus:outline-none focus:ring-2 focus:ring-indigo-black/10 focus:border-neutral-400';
  return (
    <div className="rounded-lg border border-neutral-200 bg-neutral-50/50 p-3">
      <div className="flex items-center justify-between mb-2">
        <div className="flex items-center gap-2">
          <span className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-indigo-black text-white text-[11px] font-bold">{String(idx + 1).padStart(2, '0')}</span>
          <span className="font-mono text-[11px] text-neutral-500">{item.sticker_no}</span>
        </div>
        <button type="button" onClick={onDelete} className="text-neutral-400 hover:text-rose-500"><Trash2 size={14} /></button>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-12 gap-2">
        <div className="md:col-span-12 relative">
          <input type="url" defaultValue={item.vendor_url ?? ''} placeholder="1688 상품 URL (https://detail.1688.com/offer/...)"
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
          <input type="number" defaultValue={item.unit_price || ''} placeholder="단가 ¥"
            onBlur={(e) => { const v = Number(e.target.value) || 0; if (v !== item.unit_price) onPatch({ unit_price: v }); }} className={cls} />
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
        <div className="absolute bottom-0 left-0 right-0 p-2 bg-gradient-to-t from-black/70 to-transparent">
          <div className="text-[11px] font-bold text-white truncate">{item.product_name || '(품목명 없음)'}</div>
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
