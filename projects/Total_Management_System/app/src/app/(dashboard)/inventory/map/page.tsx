'use client';

import { useState, useMemo, Fragment } from 'react';
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
import { RackEditModal } from '@/components/inventory/rack-edit-modal';
import { RackMapPrint } from '@/components/inventory/rack-map-print';
import { cellGridPos, cellSpan, levelColumnCount } from '@/lib/warehouse/location-code';
import { Modal } from '@/components/ui/modal';
import { Boxes, MapPin, Plus, ArrowLeft, Trash2, PackageX, Printer, Settings2 } from 'lucide-react';

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
  // 115: 편집·인쇄는 렉별 모달로 분리 (진입 화면은 '보기'에 집중)
  const [editRack, setEditRack] = useState<number | null>(null);
  const [printRack, setPrintRack] = useState<number | null>(null);

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

  // 렉 번호별로 묶고, 각 렉 안에서는 단(위→아래) 순. 열 수는 렉 정보(113)에서
  const racks = useMemo(() => {
    const rackInfo = new Map((data?.racks || []).map((r) => [r.rack_no, r]));
    const map = new Map<number, LocationWithProducts[]>();
    for (const l of locations) {
      const arr = map.get(l.rack_no) || [];
      arr.push(l);
      map.set(l.rack_no, arr);
    }
    return [...map.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([rackNo, list]) => {
        const info = rackInfo.get(rackNo);
        // 116: 병합 칸의 폭(col_span)까지 포함해야 오른쪽이 잘리지 않는다
        const maxBins = levelColumnCount(list);
        return {
          rackNo,
          label: info?.label || null,
          cellCount: list.length,
          // 열 수 = 렉 정보 우선, 없으면 가장 칸 많은 단 기준 (최소 1)
          columns: Math.max(1, info?.columns ?? maxBins, maxBins),
          // 115: 1단 = 맨 아래 → 화면은 큰 번호가 위로 오게 내림차순으로 그린다
          levels: [...new Set(list.map((l) => l.level_no))].sort((a, b) => b - a)
            .map((lv) => {
              const cells = list.filter((l) => l.level_no === lv)
                .sort((a, b) => ((a.bin_row ?? 0) - (b.bin_row ?? 0)) || ((a.bin_no ?? 0) - (b.bin_no ?? 0)));
              const rows = cells.reduce((m, c) => Math.max(m, c.bin_row ?? 0), 0);
              const cols = levelColumnCount(cells);
              const isShelf = cells.length === 1 && cells[0].bin_no == null;
              return {
                levelNo: lv, cells, rows, cols, isShelf,
                isDrawer: rows > 1,   // 수납함(행이 여러 개) — 단 전체를 차지하는 격자로 그린다
              };
            }),
        };
      });
  }, [locations, data]);

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

      {/* 이 칸에 제품 담기 — 한 칸에 여러 품목을 몰아 넣는 경우가 많아 여러 개를 한 번에 고른다 */}
      <AddProductToLocation
        key={selected.id}
        candidates={data?.unassigned || []}
        onAssign={(productIds) => assignLocation.mutate({ product_ids: productIds, location_id: selected.id })}
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
        <div className="space-y-5">
          {racks.map((rack) => (
            <div key={rack.rackNo}>
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-xs font-bold text-neutral-700">
                  {rack.rackNo}번 렉{rack.label ? ` · ${rack.label}` : ''}
                  <span className="ml-1.5 font-normal text-neutral-400">{rack.levels.length}단 · 자리 {rack.cellCount}개</span>
                </span>
                {/* 진입 화면은 '보기'에 집중 — 편집은 모달로, 인쇄는 여기서 바로 */}
                <div className="flex gap-1">
                  <button
                    onClick={() => setPrintRack(rack.rackNo)}
                    className="flex items-center gap-1 text-[11px] font-medium text-neutral-500 hover:text-indigo-black px-2 py-1 rounded hover:bg-neutral-100 transition"
                    title="이 렉 배치도 인쇄 (A4 한 장)"
                  ><Printer size={12} /> 인쇄</button>
                  <button
                    onClick={() => setEditRack(rack.rackNo)}
                    className="flex items-center gap-1 text-[11px] font-medium text-neutral-500 hover:text-indigo-black px-2 py-1 rounded hover:bg-neutral-100 transition"
                    title="단·열·행 수정 / 렉 삭제"
                  ><Settings2 size={12} /> 수정</button>
                </div>
              </div>

              {/* 렉 = N열 그리드. 각 단은 앞에서부터 쓰는 칸만 차지하고 나머지는 빈 공간 */}
              <div className="border-2 border-neutral-300 rounded-lg p-2 bg-neutral-50 space-y-2 overflow-x-auto">
                {rack.levels.map((lvl) => {
                  const { levelNo, cells, rows, cols, isShelf, isDrawer } = lvl;
                  const cellBtn = (loc: LocationWithProducts, compact: boolean) => {
                    const hit = matchedLocIds.has(loc.id);
                    const sel = selectedId === loc.id;
                    const empty = loc.product_count === 0;
                    const many = loc.product_count > 1;
                    const pos = cellGridPos(loc.bin_no, loc.bin_row);
                    const span = cellSpan(loc.col_span);   // 116: 가로로 병합된 넓은 칸
                    return (
                      <button
                        key={loc.id}
                        onClick={() => setSelectedId(loc.id)}
                        title={
                          empty ? `${loc.label || loc.code} · 비어 있음`
                            : `${loc.label || loc.code}\n${loc.products.map((p) => `· ${p.name} ${p.stock_quantity}개`).join('\n')}`
                        }
                        // 중간 칸을 삭제해도 나머지가 밀리지 않도록 열·행 명시 + 병합 폭(span) 반영
                        style={isShelf ? { gridColumn: '1 / -1' } : { gridColumn: `${pos.col} / span ${span}`, gridRow: pos.row }}
                        className={`rounded-md border text-left transition min-w-0 flex flex-col justify-center ${
                          compact ? 'h-[38px] px-1.5 py-1' : 'h-[70px] px-2.5 py-1.5'
                        } ${
                          sel ? 'border-neutral-900 ring-2 ring-neutral-900 bg-[#F4F0EA]'
                          : hit ? 'border-amber-500 ring-2 ring-amber-400 bg-amber-50'
                          // 빈칸 = 점선 테두리만 (요청 ③: 제품 없으면 점선 형태만)
                          : empty ? 'border-dashed border-neutral-200 bg-transparent hover:border-neutral-400 hover:bg-white'
                          : 'border-neutral-200 bg-white hover:border-neutral-400'
                        }`}
                      >
                        {/* 요청 ③: 메인 배치도는 코드(A1) 숨기고 모델명만 크게. 빈칸은 아무 글자도 안 보임 */}
                        {empty ? null : compact ? (
                          <div className="flex items-center gap-1 min-w-0">
                            <span className="text-[10px] font-bold text-indigo-black truncate leading-tight">
                              {many ? `${loc.products[0]?.name} 외` : loc.products[0]?.name}
                            </span>
                            {many && (
                              <span className="ml-auto shrink-0 px-1 rounded bg-neutral-800 text-white text-[8px] font-bold tabular-nums">{loc.product_count}</span>
                            )}
                          </div>
                        ) : (
                          <>
                            <div className="flex items-start gap-1">
                              <div className="min-w-0 flex-1">
                                {loc.products.slice(0, many ? 2 : 1).map((p) => (
                                  <div key={p.id} className="text-[13px] font-bold text-indigo-black truncate leading-tight">
                                    {p.name}
                                  </div>
                                ))}
                              </div>
                              {many && (
                                <span className="shrink-0 px-1 rounded bg-neutral-800 text-white text-[9px] font-bold tabular-nums">{loc.product_count}종</span>
                              )}
                            </div>
                            <div className="text-[10px] text-neutral-500 tabular-nums leading-tight mt-0.5">
                              {loc.product_count > 2 && <span className="text-neutral-400">외 {loc.product_count - 2}종 · </span>}
                              재고 {loc.stock_total}
                            </div>
                          </>
                        )}
                      </button>
                    );
                  };

                  return (
                    <div key={levelNo} className="flex items-stretch gap-2">
                      {isDrawer ? (
                        /* 수납함 — 단 전체를 차지하는 별도 블록 + 내부 행×열 격자 */
                        <div className="flex-1 min-w-0 rounded-md border-2 border-neutral-400 bg-white/60 p-1.5">
                          <div className="text-[10px] text-neutral-500 font-semibold mb-1">
                            {levelNo}단 · 수납함 <span className="font-normal text-neutral-400">{rows}행 {cols}열</span>
                          </div>
                          <div className="grid gap-1" style={{ gridTemplateColumns: `repeat(${cols}, minmax(52px, 1fr))` }}>
                            {cells.map((loc) => cellBtn(loc, true))}
                          </div>
                        </div>
                      ) : (
                        <div className="grid gap-2 flex-1 min-w-0"
                          style={{ gridTemplateColumns: `repeat(${rack.columns}, minmax(88px, 1fr))` }}>
                          {cells.map((loc) => cellBtn(loc, false))}
                        </div>
                      )}

                    </div>
                  );
                })}
              </div>
              <div className="text-[10px] text-neutral-400 mt-1">1단 = 맨 아래 (건물 층수와 동일) · 칸 없는 단은 선반 · 행이 여러 개면 수납함</div>
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

      {/* 115: 렉별 구조 수정 (＋열/＋행/단추가/칸삭제/렉삭제) */}
      {editRack !== null && (() => {
        const r = racks.find((x) => x.rackNo === editRack);
        if (!r) return null;
        return (
          <RackEditModal
            rackNo={r.rackNo}
            rackLabel={r.label}
            levels={r.levels}
            onClose={() => setEditRack(null)}
            onDeleteRack={() => {
              deleteLocation.mutate({ rack_no: r.rackNo });
              setEditRack(null);
              setSelectedId(null);
            }}
          />
        );
      })()}

      {/* 렉 배치도 인쇄 — A4 한 장에 렉 하나. 화면과 같은 구조(위가 큰 단)를 그대로 넘긴다 */}
      {printRack !== null && racks.some((x) => x.rackNo === printRack) && (
        <RackMapPrint
          racks={racks}
          targetRackNo={printRack}
          onClose={() => setPrintRack(null)}
        />
      )}
    </>
  );
}

/** 이 칸에 제품 담기 — 위치 미지정 제품에서 검색해 여러 개를 골라 한 번에 담는다 */
function AddProductToLocation({ candidates, onAssign, pending }: {
  candidates: LocationProduct[];
  onAssign: (productIds: string[]) => void;
  pending: boolean;
}) {
  const [open, setOpen] = useState(false);
  const [q, setQ] = useState('');
  const [picked, setPicked] = useState<Set<string>>(new Set());

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase();
    const base = s
      ? candidates.filter((p) => p.name.toLowerCase().includes(s) || (p.sku || '').toLowerCase().includes(s))
      : candidates;
    return base.slice(0, 30);
  }, [q, candidates]);

  const toggle = (id: string) => setPicked((prev) => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

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
        <span className="text-xs font-semibold text-neutral-700">제품 담기 <span className="font-normal text-neutral-400">여러 개 선택 가능</span></span>
        <button onClick={() => { setOpen(false); setQ(''); setPicked(new Set()); }}
          className="text-[11px] text-neutral-400 hover:text-neutral-600">닫기</button>
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
          filtered.map((p) => {
            const on = picked.has(p.id);
            return (
              <button
                key={p.id}
                onClick={() => toggle(p.id)}
                disabled={pending}
                className={`w-full flex items-center gap-2 px-2 py-1.5 rounded text-left transition disabled:opacity-50 ${
                  on ? 'bg-[#F4F0EA] shadow-[inset_3px_0_0_#1A1A1A]' : 'hover:bg-neutral-50'
                }`}
              >
                <input type="checkbox" checked={on} readOnly tabIndex={-1} className="w-3.5 h-3.5 shrink-0 pointer-events-none" />
                <span className="min-w-0 flex-1">
                  <span className="text-xs font-medium text-indigo-black block truncate">{p.name}</span>
                  <span className="text-[10px] text-neutral-400 font-mono">{p.sku}</span>
                </span>
                <span className="text-[11px] text-neutral-500 tabular-nums shrink-0">{p.stock_quantity}개</span>
              </button>
            );
          })
        )}
      </div>
      <button
        onClick={() => { onAssign([...picked]); setPicked(new Set()); setQ(''); setOpen(false); }}
        disabled={pending || picked.size === 0}
        className="w-full py-2 rounded-lg bg-neutral-900 text-white text-xs font-semibold transition disabled:opacity-30"
      >
        {pending ? '담는 중...' : picked.size === 0 ? '제품을 선택하세요' : `${picked.size}개 이 칸에 담기`}
      </button>
    </div>
  );
}

/** 렉 추가 — 단마다 칸 수를 다르게 지정 (실제 렉은 단별로 칸이 다르다) */
function AddRackModal({ onClose, onSubmit, pending, existingRacks }: {
  onClose: () => void;
  onSubmit: (v: { rack_no: number; label?: string | null; levels: { cols: number; rows: number }[] }) => void;
  pending: boolean;
  existingRacks: number[];
}) {
  const nextRack = existingRacks.length ? Math.max(...existingRacks) + 1 : 1;
  const [rackNo, setRackNo] = useState(nextRack);
  const [label, setLabel] = useState('');
  // 단별 {열, 행}. 행>1 이면 수납함 (예: 가위 보관함 6행 10열)
  const [levels, setLevels] = useState<{ cols: number; rows: number }[]>([
    { cols: 2, rows: 1 }, { cols: 6, rows: 1 }, { cols: 10, rows: 6 }, { cols: 0, rows: 1 }, { cols: 0, rows: 1 },
  ]);
  const dup = existingRacks.includes(rackNo);
  const simpleCols = levels.filter((l) => l.rows === 1 && l.cols > 0).map((l) => l.cols);
  const columns = Math.max(1, ...(simpleCols.length ? simpleCols : [1]));
  const totalCells = levels.reduce((s, l) => s + (l.cols > 0 ? l.cols * l.rows : 1), 0);

  const setAt = (i: number, patch: Partial<{ cols: number; rows: number }>) =>
    setLevels((prev) => prev.map((l, idx) => idx === i ? {
      cols: Math.max(0, Math.min(26, patch.cols ?? l.cols)),
      rows: Math.max(1, Math.min(20, patch.rows ?? l.rows)),
    } : l));

  return (
    // 공용 Modal — 네이티브 dialog 라 배경 클릭/드래그로 안 닫히고, preventAutoClose 로 Escape 도 막는다
    <Modal open onClose={onClose} title="렉 추가" className="max-w-md" preventAutoClose>
      <div>
        <p className="text-[11px] text-neutral-500 mb-3">단마다 열·행을 다르게 넣으세요 · 열 0 = 칸 없이 선반 · 행 2 이상 = 수납함</p>

        <div className="space-y-3">
          <div className="grid grid-cols-[100px_1fr] gap-3">
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">렉 번호</label>
              <input type="number" min={1} max={99} value={rackNo}
                onChange={(e) => setRackNo(parseInt(e.target.value) || 1)}
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
            <div>
              <label className="text-xs text-neutral-500 mb-1 block">이름 (선택)</label>
              <input type="text" value={label} onChange={(e) => setLabel(e.target.value)}
                placeholder="예: 전면 렉"
                className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm" />
            </div>
          </div>
          {dup && <p className="text-[11px] text-red-500">이미 있는 렉 번호입니다</p>}

          <div>
            <div className="flex items-center justify-between mb-1.5">
              <label className="text-xs text-neutral-500">단별 구조 <span className="text-neutral-400">(1단 = 맨 아래)</span></label>
              <div className="flex gap-1.5">
                <button onClick={() => setLevels((p) => (p.length < 20 ? [...p, { cols: 0, rows: 1 }] : p))}
                  className="text-[11px] text-neutral-500 hover:text-indigo-black underline">＋단</button>
                <button onClick={() => setLevels((p) => (p.length > 1 ? p.slice(0, -1) : p))}
                  className="text-[11px] text-neutral-400 hover:text-red-500 underline">－단</button>
              </div>
            </div>
            <div className="grid grid-cols-[3.5rem_4.5rem_4.5rem_1fr] gap-x-2 gap-y-1.5 items-center">
              <span className="text-[10px] text-neutral-400"></span>
              <span className="text-[10px] text-neutral-400">열(칸)</span>
              <span className="text-[10px] text-neutral-400">행</span>
              <span className="text-[10px] text-neutral-400">미리보기</span>
              {/* 115: 화면은 큰 번호(위)부터 — 실제 렉을 보는 순서와 맞춘다 */}
              {levels.map((_, revIdx) => levels.length - 1 - revIdx).map((i) => {
                const l = levels[i];
                return (
                <Fragment key={i}>
                  <span className="text-xs text-neutral-500">
                    {i + 1}단
                    {i === levels.length - 1 && <span className="block text-[9px] text-neutral-400 leading-none">맨 위</span>}
                    {i === 0 && <span className="block text-[9px] text-neutral-400 leading-none">맨 아래</span>}
                  </span>
                  <input type="number" min={0} max={26} value={l.cols}
                    onChange={(e) => setAt(i, { cols: parseInt(e.target.value) || 0 })}
                    className="h-8 px-2 rounded border border-neutral-200 text-sm" />
                  <input type="number" min={1} max={20} value={l.rows}
                    disabled={l.cols === 0}
                    onChange={(e) => setAt(i, { rows: parseInt(e.target.value) || 1 })}
                    className="h-8 px-2 rounded border border-neutral-200 text-sm disabled:bg-neutral-50 disabled:text-neutral-300" />
                  <div className="min-w-0">
                    {l.cols === 0 ? (
                      <div className="h-4 rounded-sm bg-neutral-300" title="칸 없이 선반" />
                    ) : l.rows > 1 ? (
                      <div className="border border-neutral-400 rounded-sm p-0.5">
                        <div className="grid gap-px" style={{ gridTemplateColumns: `repeat(${Math.min(l.cols, 12)}, 1fr)` }}>
                          {Array.from({ length: Math.min(l.cols, 12) * Math.min(l.rows, 6) }).map((_, k) => (
                            <div key={k} className="h-1.5 bg-neutral-700 rounded-[1px]" />
                          ))}
                        </div>
                      </div>
                    ) : (
                      <div className="flex gap-0.5">
                        {Array.from({ length: columns }).map((_, c) => (
                          <div key={c} className={`h-4 flex-1 rounded-sm ${c < l.cols ? 'bg-neutral-800' : 'bg-neutral-100'}`} />
                        ))}
                      </div>
                    )}
                  </div>
                  <span />
                  <span className="col-span-3 text-[10px] text-neutral-400 -mt-0.5">
                    {l.cols === 0 ? '칸 없이 선반으로 사용'
                      : l.rows > 1 ? `수납함 ${l.rows}행 ${l.cols}열 = ${l.cols * l.rows}칸`
                      : `${l.cols}칸 (한 줄)`}
                  </span>
                </Fragment>
                );
              })}
            </div>
          </div>

          <div className="p-2.5 rounded-lg bg-neutral-50 text-xs text-neutral-600">
            렉 기준 <strong className="text-indigo-black">{columns}열</strong> · 총 <strong className="text-indigo-black">{totalCells}자리</strong> 생성
            <span className="text-neutral-400 block mt-0.5">
              한 줄: R{String(rackNo).padStart(2, '0')}-1-A · 수납함: R{String(rackNo).padStart(2, '0')}-3-A1 ~ (열=알파벳, 행=숫자)
            </span>
          </div>
        </div>

        <div className="flex gap-2 pt-4 mt-3 border-t border-neutral-100">
          <Button variant="ghost" className="flex-1" onClick={onClose}>취소</Button>
          <Button className="flex-1" disabled={dup || pending}
            onClick={() => onSubmit({ rack_no: rackNo, label: label.trim() || null, levels })}>
            {pending ? '생성 중...' : '생성'}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
