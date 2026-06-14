'use client';

import { useState } from 'react';
import { ArrowLeftRight, AlertTriangle, Check, X, Search } from 'lucide-react';
import { useSerialLookup, useSwapSerials } from '@/hooks/use-serial-lookup';
import { formatPhone } from '@/lib/utils/format';

interface SerialInfo {
  id: string;
  serial_number: string;
  status: string;
  product_id: string;
  offline_sale_id: string | null;
  sold_to_name: string | null;
  sold_to_phone: string | null;
}

interface Props {
  /** 현재 시리얼 (좌측 카드) */
  currentSerial: SerialInfo;
  /** 현재 시리얼의 제품명·판매번호 (좌측 카드 표시용) */
  currentMeta: {
    product_name: string | null;
    sale_number: string | null;
  };
  /** 모달 닫기 */
  onClose: () => void;
}

/**
 * 시리얼 양방향 교환 다이얼로그 (Phase B)
 *
 * 흐름:
 *   1) 상대 시리얼 번호 입력 + Enter → lookup
 *   2) 두 시리얼 나란히 카드 비교 — 클라이언트 사전 가드 검증
 *   3) "교환합니다" 클릭 → swap_serials RPC 호출 (서버 가드 5겹 최후 검증)
 */
export function SerialSwapDialog({ currentSerial, currentMeta, onClose }: Props) {
  const [query, setQuery] = useState('');
  const [searchTrigger, setSearchTrigger] = useState('');
  const { data: lookupData, isFetching } = useSerialLookup(searchTrigger);
  const swapMutation = useSwapSerials();

  const otherSerial = lookupData?.serial;
  const otherProduct = lookupData?.product;
  const otherSale = lookupData?.sale;

  // ── 클라이언트 사전 가드 (서버 RPC가 최후 검증)
  const isSameSerial = otherSerial?.id === currentSerial.id;
  const productMatches = otherSerial && otherSerial.product_id === currentSerial.product_id;
  const sameSale = otherSerial?.offline_sale_id && currentSerial.offline_sale_id &&
                   otherSerial.offline_sale_id === currentSerial.offline_sale_id;
  const otherSold = otherSerial?.status === 'sold' && otherSerial?.offline_sale_id;

  const canSwap =
    !!otherSerial && !isSameSerial && !!productMatches && !sameSale && !!otherSold;

  const blockReason = !otherSerial
    ? null
    : isSameSerial
      ? '같은 시리얼끼리는 교환할 수 없습니다'
      : !productMatches
        ? '같은 제품의 시리얼끼리만 교환할 수 있습니다 (제품 불일치)'
        : sameSale
          ? '같은 판매 안의 시리얼끼리는 교환이 무의미합니다'
          : !otherSold
            ? '상대 시리얼이 판매완료 상태가 아닙니다'
            : null;

  function handleSearch() {
    const trimmed = query.trim();
    if (!trimmed) return;
    setSearchTrigger(trimmed);
  }

  async function handleSwap() {
    if (!otherSerial) return;
    await swapMutation.mutateAsync({
      serial_a_id: currentSerial.id,
      serial_b_id: otherSerial.id,
    });
    onClose();
  }

  return (
    <div
      className="fixed inset-0 z-[60] flex items-center justify-center bg-black/40 backdrop-blur-sm p-4"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-xl shadow-2xl w-full max-w-[640px] overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="px-5 py-4 border-b border-neutral-100 flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-full bg-neutral-100 flex items-center justify-center flex-shrink-0">
            <ArrowLeftRight size={16} className="text-neutral-700" />
          </div>
          <div className="flex-1 min-w-0">
            <h3 className="text-sm font-bold text-neutral-900 leading-tight">시리얼 교환</h3>
            <p className="text-[11px] text-neutral-500 mt-0.5">두 판매의 시리얼을 양방향으로 동시 교환합니다</p>
          </div>
          <button
            onClick={onClose}
            className="w-7 h-7 rounded-full hover:bg-neutral-100 flex items-center justify-center text-neutral-400 hover:text-neutral-700 transition"
            aria-label="닫기"
          >
            <X size={14} />
          </button>
        </div>

        {/* 검색 바 */}
        <div className="px-5 py-4 border-b border-neutral-100">
          <label className="text-[11px] font-semibold text-neutral-500 mb-1.5 block">교환할 상대 시리얼 번호</label>
          <div className="flex gap-2">
            <div className="flex-1 relative">
              <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-neutral-400" />
              <input
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && query.trim()) {
                    e.preventDefault();
                    handleSearch();
                  }
                }}
                placeholder="시리얼번호 입력 후 엔터"
                className="w-full h-9 pl-8 pr-3 rounded-lg border border-neutral-200 text-sm font-mono placeholder:text-neutral-400 placeholder:font-sans focus:outline-none focus:border-neutral-400"
                autoFocus
              />
            </div>
            <button
              type="button"
              onClick={handleSearch}
              disabled={!query.trim() || isFetching}
              className="px-3 h-9 rounded-lg bg-neutral-900 text-white text-sm font-medium disabled:opacity-30 disabled:cursor-not-allowed hover:bg-neutral-800 transition"
            >
              {isFetching ? '조회중' : '조회'}
            </button>
          </div>
        </div>

        {/* 비교 카드 영역 */}
        {searchTrigger && (
          <div className="px-5 py-4">
            {isFetching ? (
              <p className="text-center text-sm text-neutral-400 py-8">조회 중...</p>
            ) : !otherSerial ? (
              <p className="text-center text-sm text-neutral-400 py-8">
                <Search size={20} className="mx-auto mb-2 opacity-30" />
                일치하는 시리얼을 찾을 수 없습니다
              </p>
            ) : (
              <div className="grid grid-cols-[1fr_auto_1fr] gap-3 items-stretch">
                {/* 좌측 — 현재 시리얼 (A) */}
                <SwapCard
                  label="현재"
                  serial_number={currentSerial.serial_number}
                  product_name={currentMeta.product_name}
                  sale_number={currentMeta.sale_number}
                  customer_name={currentSerial.sold_to_name}
                  customer_phone={currentSerial.sold_to_phone}
                />

                {/* 가운데 화살표 */}
                <div className="flex items-center justify-center px-1">
                  <div className="w-8 h-8 rounded-full bg-neutral-900 flex items-center justify-center">
                    <ArrowLeftRight size={14} className="text-white" />
                  </div>
                </div>

                {/* 우측 — 상대 시리얼 (B) */}
                <SwapCard
                  label="상대"
                  serial_number={otherSerial.serial_number}
                  product_name={otherProduct?.name || null}
                  sale_number={otherSale?.sale_number || null}
                  customer_name={otherSerial.sold_to_name}
                  customer_phone={otherSerial.sold_to_phone}
                />
              </div>
            )}

            {/* 가드 결과 표시 */}
            {otherSerial && (
              <div className="mt-4">
                {blockReason ? (
                  <div className="flex items-start gap-2 p-3 rounded-lg bg-amber-50 border border-amber-200">
                    <AlertTriangle size={14} className="text-amber-700 shrink-0 mt-0.5" />
                    <p className="text-xs text-amber-800 leading-relaxed">{blockReason}</p>
                  </div>
                ) : (
                  <div className="flex items-start gap-2 p-3 rounded-lg bg-neutral-50 border border-neutral-200">
                    <Check size={14} className="text-neutral-700 shrink-0 mt-0.5" />
                    <p className="text-xs text-neutral-700 leading-relaxed">
                      두 시리얼은 교환 가능합니다. 실행 시 <strong>두 판매의 고객 정보·시리얼 연결이 동시에 교환</strong>됩니다.
                    </p>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {/* 푸터 */}
        <div className="px-5 py-3 border-t border-neutral-100 flex gap-2 bg-neutral-50/50">
          <button
            onClick={onClose}
            className="flex-1 h-9 rounded-lg border border-neutral-200 bg-white text-sm font-medium text-neutral-700 hover:bg-neutral-50 transition"
            disabled={swapMutation.isPending}
          >
            취소
          </button>
          <button
            onClick={handleSwap}
            disabled={!canSwap || swapMutation.isPending}
            className="flex-1 h-9 rounded-lg bg-neutral-900 text-sm font-semibold text-white hover:bg-neutral-800 transition disabled:opacity-30 disabled:cursor-not-allowed"
          >
            {swapMutation.isPending ? '교환 중...' : '교환합니다'}
          </button>
        </div>
      </div>
    </div>
  );

  function SwapCard({
    label,
    serial_number,
    product_name,
    sale_number,
    customer_name,
    customer_phone,
  }: {
    label: string;
    serial_number: string;
    product_name: string | null;
    sale_number: string | null;
    customer_name: string | null;
    customer_phone: string | null;
  }) {
    return (
      <div className="rounded-lg border border-neutral-200 bg-white p-3">
        <p className="text-[10px] font-semibold text-neutral-400 mb-1.5 uppercase tracking-wider">{label}</p>
        <p className="font-mono text-sm font-bold text-neutral-900 truncate mb-2">{serial_number}</p>
        <div className="space-y-1 text-[11px]">
          <Row label="제품" value={product_name} />
          <Row label="판매" value={sale_number} mono />
          <Row label="고객" value={customer_name} />
          <Row label="연락처" value={customer_phone ? formatPhone(customer_phone) : null} />
        </div>
      </div>
    );
  }

  function Row({ label, value, mono }: { label: string; value: string | null; mono?: boolean }) {
    return (
      <div className="flex items-baseline gap-2">
        <span className="text-neutral-400 w-12 shrink-0 whitespace-nowrap">{label}</span>
        <span className={`text-neutral-800 font-medium truncate ${mono ? 'font-mono' : ''}`}>{value || '-'}</span>
      </div>
    );
  }
}
