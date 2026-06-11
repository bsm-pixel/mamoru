'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { SearchInput } from '@/components/ui/search-input';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { useSerialLookup, useSerialAudit, type SerialAuditLog } from '@/hooks/use-serial-lookup';
import { useProducts } from '@/hooks/use-sales';
import { formatPhone } from '@/lib/utils/format';
import { Package, User, ShoppingBag, Wrench, Hash, Activity, ArrowRight, ArrowLeft, ArrowLeftRight, Search } from 'lucide-react';
import { SerialSwapDialog } from '@/components/serials/serial-swap-dialog';
import { SerialManagePanel } from '@/components/serials/serial-manage-panel';
import Link from 'next/link';

const STATUS_LABEL: Record<string, { label: string; color: string }> = {
  in_stock: { label: '재고', color: 'bg-green-100 text-green-700' },
  reserved: { label: '예약', color: 'bg-yellow-100 text-yellow-700' },
  sold: { label: '판매완료', color: 'bg-blue-100 text-blue-700' },
  returned: { label: '반품', color: 'bg-purple-100 text-purple-700' },
  defective: { label: '불량', color: 'bg-red-100 text-red-700' },
};

const ZONE_LABEL: Record<string, { label: string; color: string }> = {
  raw: { label: '보관', color: 'bg-neutral-100 text-neutral-600' },
  ready: { label: '준비', color: 'bg-green-50 text-green-700' },
  display: { label: '디스플레이', color: 'bg-blue-50 text-blue-700' },
};

const CHANNEL_LABEL: Record<string, string> = {
  offline: '오프라인',
  online: '온라인',
  talk: '온라인상담',
};

function formatDate(d: string | null) {
  if (!d) return '-';
  return new Date(d).toLocaleDateString('ko-KR', { year: 'numeric', month: 'short', day: 'numeric' });
}

function formatDateTime(d: string | null) {
  if (!d) return '-';
  return new Date(d).toLocaleString('ko-KR', {
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit',
  });
}

const ACTION_LABEL: Record<string, { label: string; color: string }> = {
  INSERT: { label: '생성', color: 'bg-green-50 text-green-700' },
  UPDATE: { label: '변경', color: 'bg-blue-50 text-blue-700' },
  DELETE: { label: '삭제', color: 'bg-red-50 text-red-700' },
};

/** audit log 한 줄을 사람이 읽기 쉬운 요약으로 변환 */
function describeChange(log: SerialAuditLog): string[] {
  const parts: string[] = [];

  if (log.action === 'INSERT') {
    const s = log.new_status ? STATUS_LABEL[log.new_status]?.label || log.new_status : null;
    const z = log.new_warehouse_zone ? ZONE_LABEL[log.new_warehouse_zone]?.label || log.new_warehouse_zone : null;
    if (s || z) parts.push([s, z].filter(Boolean).join(' / '));
    return parts;
  }

  if (log.action === 'DELETE') {
    const s = log.old_status ? STATUS_LABEL[log.old_status]?.label || log.old_status : '없음';
    parts.push(`마지막 상태: ${s}`);
    return parts;
  }

  // UPDATE — 변경 필드만 골라서 표시
  if (log.old_status !== log.new_status) {
    const from = log.old_status ? STATUS_LABEL[log.old_status]?.label || log.old_status : '없음';
    const to = log.new_status ? STATUS_LABEL[log.new_status]?.label || log.new_status : '없음';
    parts.push(`상태: ${from} → ${to}`);
  }
  if (log.old_warehouse_zone !== log.new_warehouse_zone) {
    const from = log.old_warehouse_zone ? ZONE_LABEL[log.old_warehouse_zone]?.label || log.old_warehouse_zone : '없음';
    const to = log.new_warehouse_zone ? ZONE_LABEL[log.new_warehouse_zone]?.label || log.new_warehouse_zone : '없음';
    parts.push(`위치: ${from} → ${to}`);
  }
  if (log.old_sale_item_id !== log.new_sale_item_id) {
    if (!log.old_sale_item_id && log.new_sale_item_id) parts.push('판매 연결됨');
    else if (log.old_sale_item_id && !log.new_sale_item_id) parts.push('판매 해제됨');
    else parts.push('판매 항목 교체');
  }
  if (log.old_offline_sale_id !== log.new_offline_sale_id) {
    if (!log.old_offline_sale_id && log.new_offline_sale_id) parts.push('판매 건 연결');
    else if (log.old_offline_sale_id && !log.new_offline_sale_id) parts.push('판매 건 해제');
  }
  if (log.old_contract_id !== log.new_contract_id) {
    if (!log.old_contract_id && log.new_contract_id) parts.push('계약 연결됨');
    else if (log.old_contract_id && !log.new_contract_id) parts.push('계약 해제됨');
  }
  if (log.old_product_id !== log.new_product_id) {
    parts.push('제품 변경');
  }

  return parts.length > 0 ? parts : ['변경'];
}

export default function SerialsPage() {
  const [tab, setTab] = useState<'lookup' | 'manage'>('lookup');
  const [query, setQuery] = useState('');
  const { data, isLoading, isFetched } = useSerialLookup(query);

  // 제품별 관리 탭
  const { data: products = [] } = useProducts();
  const [prodSearch, setProdSearch] = useState('');
  const [selectedProductId, setSelectedProductId] = useState<string | null>(null);

  const serial = data?.serial;
  const product = data?.product;
  const sale = data?.sale;
  const repairs = data?.repairs || [];

  // 시리얼 이동 이력 (Phase C — DB 트리거 자동 캡처)
  const { data: auditData } = useSerialAudit(serial?.id);
  const auditLogs = auditData?.logs || [];

  // 시리얼 교환 모달 (Phase B)
  const [swapOpen, setSwapOpen] = useState(false);

  return (
    <div className="flex flex-col h-full">
      <Topbar title="시리얼 관리 & 조회" />

      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {/* 탭 — 조회 / 제품별 관리 */}
        <div className="flex gap-1">
          {([['lookup', '시리얼 조회'], ['manage', '제품별 관리']] as const).map(([v, label]) => (
            <button key={v} onClick={() => setTab(v)}
              className={`px-4 py-2 rounded-lg text-xs font-semibold transition ${tab === v ? 'bg-neutral-900 text-white' : 'bg-neutral-100 text-neutral-500 hover:bg-neutral-200'}`}>
              {label}
            </button>
          ))}
        </div>

        {tab === 'lookup' && (<>
        {/* 검색 바 */}
        <SearchInput
          placeholder="시리얼번호 또는 바코드 입력..."
          value={query}
          onChange={setQuery}
        />

        {/* 로딩 */}
        {isLoading && (
          <div className="space-y-3">
            <Skeleton className="h-32 w-full" />
            <Skeleton className="h-24 w-full" />
          </div>
        )}

        {/* 결과 없음 */}
        {isFetched && !serial && query.length >= 2 && !isLoading && (
          <div className="text-center py-12 text-neutral-400">
            <Hash size={40} className="mx-auto mb-3 opacity-30" />
            <p className="text-sm">일치하는 시리얼을 찾을 수 없습니다</p>
          </div>
        )}

        {/* 결과 카드 */}
        {serial && (
          <div className="space-y-3">
            {/* 시리얼 + 상태 */}
            <div className="bg-white rounded-xl border border-neutral-200 p-4">
              <div className="flex items-center gap-2 mb-3">
                <Hash size={16} className="text-neutral-400" />
                <span className="font-mono font-bold text-sm">{serial.serial_number}</span>
              </div>
              <div className="flex flex-wrap gap-2">
                <Badge className={STATUS_LABEL[serial.status]?.color || 'bg-neutral-100'}>
                  {STATUS_LABEL[serial.status]?.label || serial.status}
                </Badge>
                <Badge className={ZONE_LABEL[serial.warehouse_zone]?.color || 'bg-neutral-100'}>
                  {ZONE_LABEL[serial.warehouse_zone]?.label || serial.warehouse_zone}
                </Badge>
                {serial.barcode && (
                  <span className="text-xs text-neutral-400">바코드: {serial.barcode}</span>
                )}
              </div>
              {serial.lot_number && (
                <p className="text-xs text-neutral-400 mt-2">로트: {serial.lot_number}</p>
              )}
            </div>

            {/* 제품 정보 */}
            {product && (
              <div className="bg-white rounded-xl border border-neutral-200 p-4">
                <div className="flex items-center gap-2 mb-2">
                  <Package size={16} className="text-neutral-400" />
                  <span className="text-xs font-semibold text-neutral-500">제품 정보</span>
                </div>
                <div className="flex items-center gap-3">
                  {product.image_url ? (
                    <img src={product.image_url} alt="" className="w-14 h-14 rounded-lg object-cover" />
                  ) : (
                    <div className="w-14 h-14 rounded-lg bg-neutral-100 flex items-center justify-center">
                      <Package size={20} className="text-neutral-300" />
                    </div>
                  )}
                  <div>
                    <p className="font-bold text-sm">{product.name}</p>
                    <p className="text-xs text-neutral-400">
                      {product.sku} · {product.category}
                    </p>
                    <p className="text-xs text-neutral-500 mt-0.5">
                      {product.price.toLocaleString()}원
                    </p>
                  </div>
                </div>
              </div>
            )}

            {/* 판매 정보 */}
            {sale && (
              <div className="bg-white rounded-xl border border-neutral-200 p-4">
                <div className="flex items-center gap-2 mb-3">
                  <ShoppingBag size={16} className="text-neutral-400" />
                  <span className="text-xs font-semibold text-neutral-500">판매 정보</span>
                  {serial.status === 'sold' && serial.offline_sale_id && (
                    <button
                      type="button"
                      onClick={() => setSwapOpen(true)}
                      className="ml-auto flex items-center gap-1 px-2 py-1 rounded-md text-[10px] font-medium bg-neutral-100 text-neutral-600 hover:bg-neutral-200 transition"
                      title="다른 시리얼과 양방향 교환"
                    >
                      <ArrowLeftRight size={10} />
                      교환
                    </button>
                  )}
                </div>
                <div className="space-y-2 text-sm">
                  <div className="flex justify-between">
                    <span className="text-neutral-400">판매번호</span>
                    <span className="font-mono">{sale.sale_number}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-neutral-400">판매일</span>
                    <span>{formatDate(sale.sale_date)}</span>
                  </div>
                  <div className="flex justify-between items-center">
                    <span className="text-neutral-400">판매경로</span>
                    <Badge className="bg-neutral-100 text-neutral-600 text-xs">
                      {CHANNEL_LABEL[sale.sale_channel] || sale.sale_channel}
                    </Badge>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-neutral-400">결제금액</span>
                    <span>{sale.total_amount.toLocaleString()}원</span>
                  </div>
                </div>
              </div>
            )}

            {/* 고객 정보 */}
            {(serial.sold_to_name || sale?.customer_name) && (
              <div className="bg-white rounded-xl border border-neutral-200 p-4">
                <div className="flex items-center gap-2 mb-3">
                  <User size={16} className="text-neutral-400" />
                  <span className="text-xs font-semibold text-neutral-500">고객 정보</span>
                </div>
                <div className="space-y-2 text-sm">
                  <div className="flex justify-between">
                    <span className="text-neutral-400">성함</span>
                    <span className="font-medium">{serial.sold_to_name || sale?.customer_name}</span>
                  </div>
                  {(serial.sold_to_phone || sale?.customer_phone) && (
                    <div className="flex justify-between">
                      <span className="text-neutral-400">연락처</span>
                      <span>{formatPhone(serial.sold_to_phone || sale?.customer_phone || null)}</span>
                    </div>
                  )}
                </div>
              </div>
            )}

            {/* 미판매 안내 */}
            {serial.status !== 'sold' && (
              <div className="bg-neutral-50 rounded-xl border border-neutral-200 p-4 text-center">
                <p className="text-sm text-neutral-500">아직 판매되지 않은 제품입니다</p>
              </div>
            )}

            {/* 복원수리 이력 */}
            {repairs.length > 0 && (
              <div className="bg-white rounded-xl border border-neutral-200 p-4">
                <div className="flex items-center gap-2 mb-3">
                  <Wrench size={16} className="text-neutral-400" />
                  <span className="text-xs font-semibold text-neutral-500">복원수리 이력</span>
                </div>
                <div className="space-y-2">
                  {repairs.map((r) => (
                    <Link
                      key={r.id}
                      href={`/repairs/${r.id}`}
                      className="flex justify-between items-center p-2 rounded-lg hover:bg-neutral-50 transition"
                    >
                      <span className="text-sm text-neutral-600">{formatDate(r.created_at)}</span>
                      <Badge className="bg-neutral-100 text-neutral-600 text-xs">{r.status}</Badge>
                    </Link>
                  ))}
                </div>
              </div>
            )}

            {/* 이동 이력 (Phase C — append-only ledger) */}
            {auditLogs.length > 0 && (
              <div className="bg-white rounded-xl border border-neutral-200 p-4">
                <div className="flex items-center gap-2 mb-3">
                  <Activity size={16} className="text-neutral-400" />
                  <span className="text-xs font-semibold text-neutral-500">이동 이력</span>
                  <span className="text-[10px] text-neutral-300 ml-auto">최신순 · 최대 100건</span>
                </div>
                <div className="space-y-2">
                  {auditLogs.map((log) => {
                    const action = ACTION_LABEL[log.action] || ACTION_LABEL.UPDATE;
                    const changes = describeChange(log);
                    return (
                      <div key={log.id} className="flex items-start gap-2 p-2 rounded-lg hover:bg-neutral-50 transition">
                        <Badge className={`${action.color} text-[10px] shrink-0 mt-0.5`}>{action.label}</Badge>
                        <div className="flex-1 min-w-0">
                          <div className="flex flex-wrap items-center gap-1 text-sm">
                            {changes.map((c, i) => (
                              <span key={i} className="text-neutral-700">
                                {i > 0 && <ArrowRight size={11} className="inline mx-1 text-neutral-300" />}
                                {c}
                              </span>
                            ))}
                          </div>
                          <p className="text-[11px] text-neutral-400 mt-0.5">
                            {formatDateTime(log.changed_at)} · {log.changed_by_name}
                          </p>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* 초기 안내 */}
        {!query && (
          <div className="text-center py-16 text-neutral-300">
            <Hash size={48} className="mx-auto mb-4 opacity-30" />
            <p className="text-sm">시리얼번호 또는 바코드를 입력하세요</p>
            <p className="text-xs text-neutral-300 mt-1">제품 정보, 구매자, 판매경로를 즉시 확인합니다</p>
          </div>
        )}
        </>)}

        {/* 제품별 관리 탭 — 제품 선택 → 시리얼 생성·상태/위치 관리 */}
        {tab === 'manage' && (
          selectedProductId ? (
            <div className="space-y-3">
              <button onClick={() => setSelectedProductId(null)} className="text-xs text-blue-600 hover:underline flex items-center gap-1">
                <ArrowLeft size={12} /> 다른 제품 선택
              </button>
              {(() => { const p = products.find((x) => x.id === selectedProductId); return p ? (
                <div className="flex items-center gap-2">
                  <Package size={16} className="text-stone-900" />
                  <span className="text-sm font-bold">{p.name}</span>
                  <span className="text-xs text-neutral-500">({p.sku})</span>
                </div>
              ) : null; })()}
              <SerialManagePanel key={selectedProductId} productId={selectedProductId} productName={products.find((x) => x.id === selectedProductId)?.name} />
            </div>
          ) : (
            <div className="space-y-3">
              <div className="relative">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
                <input type="text" value={prodSearch} onChange={(e) => setProdSearch(e.target.value)} placeholder="제품명 또는 SKU 검색"
                  className="w-full h-10 pl-9 pr-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-stone-400" />
              </div>
              <div className="bg-white rounded-xl border border-neutral-200 divide-y divide-neutral-100 max-h-[60vh] overflow-y-auto">
                {products.filter((p) => p.category !== 'SUP' && (!prodSearch || p.name.toLowerCase().includes(prodSearch.toLowerCase()) || (p.sku || '').toLowerCase().includes(prodSearch.toLowerCase()))).map((p) => (
                  <button key={p.id} onClick={() => setSelectedProductId(p.id)} className="w-full flex items-center justify-between px-4 py-3 text-left hover:bg-neutral-50 transition">
                    <div className="flex items-center gap-2 min-w-0">
                      <Package size={15} className="text-neutral-400 shrink-0" />
                      <span className="text-sm font-medium truncate">{p.name}</span>
                      <span className="text-xs text-neutral-400 shrink-0">{p.sku}</span>
                    </div>
                    <ArrowRight size={14} className="text-neutral-300 shrink-0" />
                  </button>
                ))}
                {products.filter((p) => p.category !== 'SUP').length === 0 && (
                  <div className="px-4 py-8 text-center text-sm text-neutral-400">제품이 없습니다</div>
                )}
              </div>
            </div>
          )
        )}
      </div>

      {/* 시리얼 교환 모달 (Phase B) */}
      {swapOpen && serial && (
        <SerialSwapDialog
          currentSerial={{
            id: serial.id,
            serial_number: serial.serial_number,
            status: serial.status,
            product_id: serial.product_id,
            offline_sale_id: serial.offline_sale_id,
            sold_to_name: serial.sold_to_name,
            sold_to_phone: serial.sold_to_phone,
          }}
          currentMeta={{
            product_name: product?.name || null,
            sale_number: sale?.sale_number || null,
          }}
          onClose={() => setSwapOpen(false)}
        />
      )}
    </div>
  );
}
