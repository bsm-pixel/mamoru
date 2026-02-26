'use client';

import { use, useState } from 'react';
import { useRouter } from 'next/navigation';
import { Topbar } from '@/components/layout/topbar';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Skeleton } from '@/components/ui/skeleton';
import { useSerials, useCreateSerial, useCreateSerialBatch, useUpdateSerialStatus } from '@/hooks/use-serials';
import { useProducts } from '@/hooks/use-sales';
import { formatDateTime } from '@/lib/utils/format';
import { ArrowLeft, Plus, Package, Search, Barcode } from 'lucide-react';

const STATUS_COLOR: Record<string, string> = {
  in_stock: 'bg-green-100 text-green-700',
  reserved: 'bg-yellow-100 text-yellow-700',
  sold: 'bg-blue-100 text-blue-700',
  returned: 'bg-purple-100 text-purple-700',
  defective: 'bg-red-100 text-red-700',
};

const STATUS_LABEL: Record<string, string> = {
  in_stock: '재고',
  reserved: '예약',
  sold: '판매',
  returned: '반품',
  defective: '불량',
};

const STATUS_TABS = [
  { value: 'all', label: '전체' },
  { value: 'in_stock', label: '재고' },
  { value: 'sold', label: '판매' },
  { value: 'reserved', label: '예약' },
  { value: 'returned', label: '반품' },
  { value: 'defective', label: '불량' },
];

export default function SerialsPage({ params }: { params: Promise<{ id: string }> }) {
  const { id: productId } = use(params);
  const router = useRouter();
  const [status, setStatus] = useState('all');
  const [search, setSearch] = useState('');
  const [showAddForm, setShowAddForm] = useState(false);
  const [showBatchForm, setShowBatchForm] = useState(false);
  const [newSerial, setNewSerial] = useState('');
  const [newBarcode, setNewBarcode] = useState('');
  const [batchCount, setBatchCount] = useState(10);
  const [batchLot, setBatchLot] = useState('');

  const { data: products = [] } = useProducts();
  const product = products.find((p) => p.id === productId);

  const { data, isLoading } = useSerials(productId, { status, search, limit: 50 });
  const serials = data?.serials || [];
  const total = data?.total || 0;

  const createSerial = useCreateSerial();
  const createBatch = useCreateSerialBatch();
  const updateStatus = useUpdateSerialStatus();

  async function handleAddSerial() {
    if (!newSerial.trim()) return;
    await createSerial.mutateAsync({
      product_id: productId,
      serial_number: newSerial.trim(),
      barcode: newBarcode.trim() || undefined,
    });
    setNewSerial('');
    setNewBarcode('');
    setShowAddForm(false);
  }

  async function handleBatch() {
    if (batchCount < 1) return;
    await createBatch.mutateAsync({
      product_id: productId,
      count: batchCount,
      lot_number: batchLot.trim() || undefined,
    });
    setShowBatchForm(false);
  }

  return (
    <>
      <Topbar title={product ? `${product.name} — 시리얼 관리` : '시리얼 관리'} />

      <div className="px-4 md:px-6 py-4 space-y-4">
        <div className="flex items-center gap-3">
          <Button variant="ghost" size="sm" onClick={() => router.push('/products')}>
            <ArrowLeft size={14} />
          </Button>

          {product && (
            <div className="flex items-center gap-2">
              <Package size={16} className="text-terracotta" />
              <span className="text-sm font-bold">{product.name}</span>
              <span className="text-xs text-neutral-500">({product.sku})</span>
              <Badge className="bg-neutral-100 text-neutral-600">
                총 {total}개
              </Badge>
            </div>
          )}
        </div>

        {/* 액션 바 */}
        <div className="flex items-center gap-2 flex-wrap">
          <Button size="sm" onClick={() => { setShowAddForm(!showAddForm); setShowBatchForm(false); }}>
            <Plus size={14} />
            단건 등록
          </Button>
          <Button variant="secondary" size="sm" onClick={() => { setShowBatchForm(!showBatchForm); setShowAddForm(false); }}>
            <Barcode size={14} />
            일괄 생성
          </Button>

          <div className="flex-1 relative">
            <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-neutral-400" />
            <input
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="시리얼번호, 바코드, 고객명 검색"
              className="w-full h-9 pl-9 pr-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm text-indigo-black placeholder:text-neutral-400 focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
            />
          </div>
        </div>

        {/* 단건 등록 폼 */}
        {showAddForm && (
          <Card>
            <h4 className="text-sm font-semibold mb-2">시리얼 단건 등록</h4>
            <div className="flex gap-2">
              <input type="text" value={newSerial} onChange={(e) => setNewSerial(e.target.value)} placeholder="시리얼번호 *" className="flex-1 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
              <input type="text" value={newBarcode} onChange={(e) => setNewBarcode(e.target.value)} placeholder="바코드" className="w-40 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
              <Button size="sm" onClick={handleAddSerial} disabled={!newSerial.trim() || createSerial.isPending}>
                {createSerial.isPending ? '...' : '등록'}
              </Button>
            </div>
          </Card>
        )}

        {/* 일괄 생성 폼 */}
        {showBatchForm && (
          <Card>
            <h4 className="text-sm font-semibold mb-2">일괄 생성 (자동 넘버링)</h4>
            <div className="flex gap-2 items-end">
              <div>
                <label className="text-xs text-neutral-500">수량 (1~100)</label>
                <input type="number" value={batchCount} onChange={(e) => setBatchCount(parseInt(e.target.value) || 0)} min={1} max={100} className="w-20 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
              </div>
              <div>
                <label className="text-xs text-neutral-500">로트번호</label>
                <input type="text" value={batchLot} onChange={(e) => setBatchLot(e.target.value)} placeholder="LOT-001" className="w-32 h-9 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-sm" />
              </div>
              <Button size="sm" onClick={handleBatch} disabled={batchCount < 1 || createBatch.isPending}>
                {createBatch.isPending ? '...' : `${batchCount}개 생성`}
              </Button>
            </div>
            <p className="text-xs text-neutral-400 mt-2">
              형식: MM-{product?.sku || 'SKU'}-{new Date().toISOString().slice(0, 10).replace(/-/g, '')}-001
            </p>
          </Card>
        )}

        {/* 상태 탭 */}
        <div className="flex gap-1 overflow-x-auto pb-1">
          {STATUS_TABS.map((tab) => (
            <button
              key={tab.value}
              onClick={() => setStatus(tab.value)}
              className={`px-3 py-1.5 rounded-lg text-xs font-semibold whitespace-nowrap transition ${
                status === tab.value
                  ? 'bg-terracotta text-cream'
                  : 'bg-card-white text-neutral-500 hover:bg-warm-ivory'
              }`}
            >
              {tab.label}
            </button>
          ))}
        </div>

        {/* 시리얼 목록 */}
        <Card padding={false}>
          {isLoading ? (
            <div className="p-4 space-y-2">
              {Array.from({ length: 5 }).map((_, i) => (
                <Skeleton key={i} className="h-12 w-full" />
              ))}
            </div>
          ) : serials.length === 0 ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">
              시리얼 정보가 없습니다
            </div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {serials.map((serial) => (
                <div key={serial.id} className="flex items-center gap-3 px-4 py-2.5">
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2">
                      <span className="text-sm font-mono font-semibold text-indigo-black">
                        {serial.serial_number}
                      </span>
                      <Badge className={STATUS_COLOR[serial.status] || ''}>
                        {STATUS_LABEL[serial.status] || serial.status}
                      </Badge>
                    </div>
                    <div className="flex items-center gap-3 mt-0.5 text-xs text-neutral-500">
                      {serial.barcode && <span className="font-mono">{serial.barcode}</span>}
                      {serial.lot_number && <span>LOT: {serial.lot_number}</span>}
                      {serial.sold_to_name && <span className="text-blue-600">{serial.sold_to_name}</span>}
                      <span>{formatDateTime(serial.created_at)}</span>
                    </div>
                  </div>

                  {/* 상태 변경 드롭다운 */}
                  {serial.status === 'in_stock' && (
                    <select
                      className="text-xs border border-neutral-200 rounded px-2 py-1 bg-white"
                      value=""
                      onChange={(e) => {
                        if (e.target.value) {
                          updateStatus.mutate({ id: serial.id, status: e.target.value });
                        }
                      }}
                    >
                      <option value="">변경</option>
                      <option value="reserved">예약</option>
                      <option value="defective">불량</option>
                    </select>
                  )}
                </div>
              ))}
            </div>
          )}
        </Card>
      </div>
    </>
  );
}
