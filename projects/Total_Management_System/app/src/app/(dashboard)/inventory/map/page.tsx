'use client';

import { useState, useMemo } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { SearchInput } from '@/components/ui/search-input';
import { EmptyState } from '@/components/ui/empty-state';
import { StatCard } from '@/components/ui/stat-card';
import { SlidePanel } from '@/components/ui/slide-panel';
import { useIsLg } from '@/hooks/use-grid-mode';
import { useLocations, useCreateRack, useDeleteLocation, useAssignLocation, type LocationWithProducts, type LocationProduct } from '@/hooks/use-warehouse';
import { Boxes, MapPin, Plus, ArrowLeft, Trash2, PackageX } from 'lucide-react';

/**
 * 창고 배치도 (정위치 관리) — 112, 2026-07-18
 * 렉을 그림으로 배치하고, 각 칸에 어떤 제품이 있는지 + 재고를 함께 본다.
 * 상단 검색 → 해당 제품이 있는 칸을 하이라이트 ("이거 어디 있지?" 해결)
 */
export default function WarehouseLayoutPage() {
  const router = useRouter();
  const isLg = useIsLg();
  const { data, isLoading } = useLocations();
  const createRack = useCreateRack();
  const deleteLocation = useDeleteLocation();
  const assignLocation = useAssignLocation();

  const [search, setSearch] = useState('');
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [showAddRack, setShowAddRack] = useState(false);

  const locations = useMemo(() => data?.locations || [], [data]);

  // 검색어와 일치하는 제품이 있는 칸 → 하이라이트 대상
  const matchedLocIds = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return new Set<string>();
    const s = new Set<string>();
    for (const loc of locations) {
      if (loc.products.some((p) => p.name.toLowerCase().includes(q) || (p.sku || '').toLowerCase().includes(q))) {
        s.add(loc.id);
      }
    }
    return s;
  }, [search, locations]);

  // 렉 번호별로 묶고, 각 렉 안에서는 단(위→아래) 순
  const racks = useMemo(() => {
    const map = new Map<number, LocationWithProducts[]>();
    for (const l of locations) {
      const arr = map.get(l.rack_no) || [];
      arr.push(l);
      map.set(l.rack_no, arr);
    }
    return [...map.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([rackNo, list]) => ({
        rackNo,
        levels: [...new Set(list.map((l) => l.level_no))].sort((a, b) => a - b)
          .map((lv) => ({ levelNo: lv, bins: list.filter((l) => l.level_no === lv).sort((a, b) => (a.bin_no ?? 0) - (b.bin_no ?? 0)) })),
      }));
  }, [locations]);

  const selected = locations.find((l) => l.id === selectedId) || null;
  const totalStock = locations.reduce((s, l) => s + l.stock_total, 0);
  const emptyBins = locations.filter((l) => l.product_count === 0).length;

  const detailPanel = selected ? (
    <div className="p-4 space-y-3">
      <div>
        <div className="flex items-center gap-2">
          <span className="px-2 py-0.5 rounded bg-neutral-900 text-white text-xs font-mono font-bold">{selected.code}</span>
          <h3 className="text-base font-bold text-indigo-black">{selected.label || selected.code}</h3>
        </div>
        <p className="text-xs text-neutral-500 mt-1">
          제품 {selected.product_count}종 · 재고 합 {selected.stock_total}개
        </p>
      </div>

      {selected.products.length === 0 ? (
        <div className="py-10 text-center text-sm text-neutral-400">
          <PackageX size={24} className="mx-auto mb-2 opacity-40" />
          이 칸은 비어 있습니다
          <p className="text-xs mt-1 text-neutral-400">재고 목록에서 제품에 이 위치를 지정하세요</p>
        </div>
      ) : (
        <div className="divide-y divide-neutral-100 border border-neutral-200 rounded-lg overflow-hidden">
          {selected.products.map((p) => (
            <div key={p.id} className="flex items-center gap-3 px-3 py-2.5 hover:bg-warm-ivory/50 transition">
              <div className="flex-1 min-w-0">
                <p className="text-sm font-semibold text-indigo-black truncate">{p.name}</p>
                <p className="text-[11px] text-neutral-400 font-mono">{p.sku}</p>
              </div>
              <span className="text-sm font-bold tabular-nums shrink-0">{p.stock_quantity}개</span>
              <button
                onClick={() => assignLocation.mutate({ product_ids: [p.id], location_id: null })}
                disabled={assignLocation.isPending}
                className="text-[11px] text-neutral-400 hover:text-red-500 underline shrink-0 disabled:opacity-50"
                title="이 칸에서 빼기 (위치 미지정으로)"
              >
                빼기
              </button>
            </div>
          ))}
        </div>
      )}

      {/* 이 칸에 제품 담기 — 위치 미지정 제품에서 고른다 */}
      <AddProductToLocation
        locationId={selected.id}
        candidates={data?.unassigned || []}
        onAssign={(productId) => assignLocation.mutate({ product_ids: [productId], location_id: selected.id })}
        pending={assignLocation.isPending}
      />

      <button
        onClick={() => {
          if (!window.confirm(`'${selected.label || selected.code}' 칸을 삭제할까요?\n배정된 제품 ${selected.product_count}종은 '위치 미지정'으로 바뀝니다.`)) return;
          deleteLocation.mutate({ id: selected.id });
          setSelectedId(null);
        }}
        disabled={deleteLocation.isPending}
        className="w-full flex items-center justify-center gap-1.5 py-2 text-xs text-neutral-400 hover:text-red-600 transition disabled:opacity-50"
      >
        <Trash2 size={12} /> 이 칸 삭제
      </button>
    </div>
  ) : (
    <div className="flex flex-col items-center justify-center h-60 text-neutral-400">
      <MapPin size={28} className="mb-2 opacity-40" />
      <p className="text-xs">배치도에서 칸을 선택하세요</p>
    </div>
  );

  const mapContent = (
    <div className="space-y-4">
      {/* 요약 */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <StatCard compact label="렉" icon={Boxes} value={racks.length} primarySub="등록됨" />
        <StatCard compact label="칸" icon={MapPin} value={data?.total_locations || 0} primarySub={`빈칸 ${emptyBins}`} />
        <StatCard compact label="배치 재고" icon={Boxes} accent="emerald" value={totalStock} primarySub="개" />
        <StatCard compact label="위치 미지정" icon={PackageX} accent="amber" dimWhenZero
          value={data?.unassigned_count || 0} primarySub="제품" />
      </div>

      {/* 검색 — 제품이 어느 칸에 있는지 찾기 */}
      <div className="flex items-center gap-2">
        <SearchInput value={search} onChange={setSearch} placeholder="제품명·SKU 검색 → 있는 칸이 표시됩니다" />
        <Button size="sm" variant="secondary" onClick={() => setShowAddRack(true)}>
          <Plus size={14} /> 렉 추가
        </Button>
      </div>
      {search.trim() && (
        <p className="text-xs text-neutral-500">
          {matchedLocIds.size > 0
            ? <>🔦 <strong className="text-indigo-black">{matchedLocIds.size}개 칸</strong>에서 찾았습니다 — 아래 강조된 칸을 보세요</>
            : <span className="text-amber-600">해당 제품이 배치된 칸이 없습니다 (위치 미지정일 수 있어요)</span>}
        </p>
      )}

      {/* 렉 배치도 */}
      {racks.length === 0 ? (
        <EmptyState icon={Boxes} message="등록된 렉이 없습니다 — [렉 추가]로 시작하세요" />
      ) : (
        <div className="flex gap-4 overflow-x-auto pb-2">
          {racks.map((rack) => (
            <div key={rack.rackNo} className="shrink-0">
              <div className="text-center mb-1.5">
                <span className="text-xs font-bold text-neutral-700">{rack.rackNo}번 렉</span>
              </div>
              <div className="border-2 border-neutral-300 rounded-lg p-1.5 bg-neutral-50 space-y-1.5">
                {rack.levels.map(({ levelNo, bins }) => (
                  <div key={levelNo} className="flex gap-1.5">
                    {bins.map((loc) => {
                      const hit = matchedLocIds.has(loc.id);
                      const sel = selectedId === loc.id;
                      const empty = loc.product_count === 0;
                      return (
                        <button
                          key={loc.id}
                          onClick={() => setSelectedId(loc.id)}
                          title={`${loc.label || loc.code} · 제품 ${loc.product_count}종 / 재고 ${loc.stock_total}개`}
                          className={`w-[92px] h-[62px] rounded-md border text-left px-2 py-1.5 transition ${
                            sel ? 'border-neutral-900 ring-2 ring-neutral-900 bg-[#F4F0EA]'
                            : hit ? 'border-amber-500 ring-2 ring-amber-400 bg-amber-50'
                            : empty ? 'border-dashed border-neutral-300 bg-white hover:border-neutral-400'
                            : 'border-neutral-200 bg-white hover:border-neutral-400'
                          }`}
                        >
                          <div className="text-[10px] font-mono text-neutral-400 truncate">{loc.code}</div>
                          {empty ? (
                            <div className="text-[11px] text-neutral-300 mt-1">비어 있음</div>
                          ) : (
                            <>
                              <div className="text-[11px] font-semibold text-indigo-black truncate mt-0.5">
                                {loc.products[0]?.name}{loc.product_count > 1 ? ` 외 ${loc.product_count - 1}` : ''}
                              </div>
                              <div className="text-[10px] text-neutral-500 tabular-nums">재고 {loc.stock_total}</div>
                            </>
                          )}
                        </button>
                      );
                    })}
                  </div>
                ))}
              </div>
              <div className="text-center mt-1">
                <span className="text-[10px] text-neutral-400">위 → 아래 (1단=상단)</span>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );

  return (
    <>
      <Topbar title="창고 배치도" action={
        <Button size="sm" variant="ghost" onClick={() => router.push('/inventory')}>
          <ArrowLeft size={14} /> 재고 목록
        </Button>
      } />

      <div className="bg-stone-50 min-h-screen px-4 md:px-6 py-4">
        {isLoading ? (
          <div className="space-y-3">
            {Array.from({ length: 3 }).map((_, i) => <Skeleton key={i} className="h-24 w-full" />)}
          </div>
        ) : isLg ? (
          <div className="flex gap-4">
            <div className="flex-1 min-w-0">{mapContent}</div>
            <div className="w-[400px] shrink-0">
              <Card padding={false}>{detailPanel}</Card>
            </div>
          </div>
        ) : (
          <>
            {mapContent}
            <SlidePanel open={!!selectedId} onClose={() => setSelectedId(null)} title="칸 상세">
              {detailPanel}
            </SlidePanel>
          </>
        )}
      </div>

      {showAddRack && (
        <AddRackModal
          onClose={() => setShowAddRack(false)}
          onSubmit={(v) => { createRack.mutate(v, { onSuccess: () => setShowAddRack(false) }); }}
          pending={createRack.isPending}
          existingRacks={racks.map((r) => r.rackNo)}
        />
      )}
    </>
  );
}

/** 이 칸에 제품 담기 — 위치 미지정 제품 중에서 검색해 고른다 */
function AddProductToLocation({ candidates, onAssign, pending }: {
  locationId: string;
  candidates: LocationProduct[];
  onAssign: (productId: string) => void;
  pending: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    const base = s
      ? candidates.filter((p) => p.name.toLowerCase().includes(s) || (p.sku || '').toLowerCase().includes(s))
      : candidates;
    return base.slice(0, 20);
  }, [q, candidates]);

  if (!open) {
    return (
      <button
        onClick={() => setOpen(true)}
        className="w-full flex items-center justify-center gap-1.5 py-2 rounded-lg border border-dashed border-neutral-300 text-xs font-semibold text-neutral-600 hover:bg-neutral-50 transition"
      >
        <Plus size={13} /> 이 칸에 제품 담기
        {candidates.length > 0 && <span className="text-neutral-400">(미지정 {candidates.length})</span>}
      </button>
    );
  }

  return (
    <div className="rounded-lg border border-neutral-200 p-2.5 space-y-2">
      <div className="flex items-center justify-between">
        <span className="text-xs font-semibold text-neutral-700">제품 담기</span>
        <button onClick={() => { setOpen(false); setQ(''); }} className="text-[11px] text-neutral-400 hover:text-neutral-600">닫기</button>
      </div>
      <input
        autoFocus
        value={q}
        onChange={(e) => setQ(e.target.value)}
        placeholder="제품명·SKU 검색"
        className="w-full h-8 px-2 rounded border border-neutral-200 text-xs focus:outline-none focus:ring-2 focus:ring-neutral-300"
      />
      <div className="max-h-48 overflow-y-auto space-y-0.5">
        {candidates.length === 0 ? (
          <p className="text-[11px] text-neutral-400 text-center py-4">위치 미지정 제품이 없습니다</p>
        ) : filtered.length === 0 ? (
          <p className="text-[11px] text-neutral-400 text-center py-4">검색 결과 없음</p>
        ) : (
          filtered.map((p) => (
            <button
              key={p.id}
              onClick={() => { onAssign(p.id); setQ(''); }}
              disabled={pending}
              className="w-full flex items-center justify-between gap-2 px-2 py-1.5 rounded hover:bg-neutral-50 text-left transition disabled:opacity-50"
            >
              <span className="min-w-0">
                <span className="text-xs font-medium text-indigo-black block truncate">{p.name}</span>
                <span className="text-[10px] text-neutral-400 font-mono">{p.sku}</span>
              </span>
              <span className="text-[11px] text-neutral-500 tabular-nums shrink-0">{p.stock_quantity}개</span>
            </button>
          ))
        )}
      </div>
    </div>
  );
}

/** 렉 추가 — 단/칸 수를 넣으면 칸이 자동 생성된다 */
function AddRackModal({ onClose, onSubmit, pending, existingRacks }: {
  onClose: () => void;
  onSubmit: (v: { rack_no: number; levels: number; bins?: number }) => void;
  pending: boolean;
  existingRacks: number[];
}) {
  const nextRack = existingRacks.length ? Math.max(...existingRacks) + 1 : 1;
  const [rackNo, setRackNo] = useState(nextRack);
  const [levels, setLevels] = useState(3);
  const [bins, setBins] = useState(0);
  const dup = existingRacks.includes(rackNo);
  const preview = levels * (bins > 0 ? bins : 1);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl w-full max-w-sm mx-4 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-neutral-100">
          <h2 className="text-sm font-bold text-indigo-black">렉 추가</h2>
          <p className="text-[11px] text-neutral-500 mt-0.5">단·칸 수를 넣으면 자리가 자동으로 만들어집니다</p>
        </div>
        <div className="px-5 py-4 space-y-3">
          <div>
            <label className="text-xs text-neutral-500 mb-1 block">렉 번호</label>
            <input type="number" min={1} max={99} value={rackNo}
              onChange={(e) => setRackNo(parseInt(e.target.value) || 1)}
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            {dup && <p className="text-[11px] text-red-500 mt-1">이미 있는 렉 번호입니다</p>}
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">단 수</label>
              <input type="number" min={1} max={20} value={levels}
                onChange={(e) => setLevels(parseInt(e.target.value) || 1)}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              <p className="text-[10px] text-neutral-400 mt-0.5">3단이면 상·중·하</p>
            </div>
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">칸 수 (선택)</label>
              <input type="number" min={0} max={26} value={bins}
                onChange={(e) => setBins(parseInt(e.target.value) || 0)}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
              <p className="text-[10px] text-neutral-400 mt-0.5">0 = 칸 안 나눔</p>
            </div>
          </div>
          <div className="p-2.5 rounded-lg bg-neutral-50 text-xs text-neutral-600">
            총 <strong className="text-indigo-black">{preview}개</strong> 자리가 생성됩니다
            <span className="text-neutral-400"> · 예: {`R${String(rackNo).padStart(2, '0')}-1${bins > 0 ? '-A' : ''}`}</span>
          </div>
        </div>
        <div className="px-5 py-4 border-t border-neutral-100 flex gap-2">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={dup || pending}
            onClick={() => onSubmit({ rack_no: rackNo, levels, bins })}>
            {pending ? '생성 중...' : '생성'}
          </Button>
        </div>
      </div>
    </div>
  );
}
