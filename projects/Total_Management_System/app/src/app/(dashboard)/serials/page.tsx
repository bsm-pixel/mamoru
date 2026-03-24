'use client';

import { useState } from 'react';
import { Topbar } from '@/components/layout/topbar';
import { SearchInput } from '@/components/ui/search-input';
import { Badge } from '@/components/ui/badge';
import { Skeleton } from '@/components/ui/skeleton';
import { useSerialLookup } from '@/hooks/use-serial-lookup';
import { formatPhone } from '@/lib/utils/format';
import { Package, User, ShoppingBag, Wrench, Hash } from 'lucide-react';
import Link from 'next/link';

const STATUS_LABEL: Record<string, { label: string; color: string }> = {
  in_stock: { label: '재고', color: 'bg-green-100 text-green-700' },
  reserved: { label: '예약', color: 'bg-yellow-100 text-yellow-700' },
  sold: { label: '판매완료', color: 'bg-blue-100 text-blue-700' },
  returned: { label: '반품', color: 'bg-purple-100 text-purple-700' },
  defective: { label: '불량', color: 'bg-red-100 text-red-700' },
};

const ZONE_LABEL: Record<string, { label: string; color: string }> = {
  raw: { label: '매입원본', color: 'bg-neutral-100 text-neutral-600' },
  ready: { label: '판매준비', color: 'bg-green-50 text-green-700' },
  display: { label: '진열', color: 'bg-blue-50 text-blue-700' },
};

const CHANNEL_LABEL: Record<string, string> = {
  offline: '오프라인',
  online: '온라인',
  talk: '톡상담',
};

function formatDate(d: string | null) {
  if (!d) return '-';
  return new Date(d).toLocaleDateString('ko-KR', { year: 'numeric', month: 'short', day: 'numeric' });
}

export default function SerialsPage() {
  const [query, setQuery] = useState('');
  const { data, isLoading, isFetched } = useSerialLookup(query);

  const serial = data?.serial;
  const product = data?.product;
  const sale = data?.sale;
  const repairs = data?.repairs || [];

  return (
    <div className="flex flex-col h-full">
      <Topbar title="시리얼 조회" />

      <div className="flex-1 overflow-y-auto p-4 space-y-4">
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
      </div>
    </div>
  );
}
