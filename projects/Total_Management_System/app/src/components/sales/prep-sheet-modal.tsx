'use client';

import { useState, useRef, useEffect } from 'react';
import { Printer, X } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';

interface SaleData {
  sale: {
    id: string;
    sale_number: string;
    sale_date: string;
    customer_name: string;
    customer_phone?: string | null;
    memo?: string | null;
    sale_channel?: string | null;
    payment_status?: string | null;
    [key: string]: unknown;
  };
  items: Array<{
    id: string;
    product_id?: string | null;
    product_name: string;
    sku?: string | null;
    quantity: number;
    unit_price: number;
    [key: string]: unknown;
  }>;
  serials: Array<{
    id: string;
    serial_number: string;
    product_id?: string | null;
    sale_item_id?: string | null;
    [key: string]: unknown;
  }>;
}

interface PrepSheetModalProps {
  saleIds: string[];
  /** 단건 모드: 이미 로드된 데이터 직접 전달 */
  preloaded?: SaleData;
  onClose: () => void;
}

export function PrepSheetModal({ saleIds, preloaded, onClose }: PrepSheetModalProps) {
  const printRef = useRef<HTMLDivElement>(null);
  const [salesData, setSalesData] = useState<SaleData[]>([]);
  const [memos, setMemos] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(false);

  // 데이터 로딩
  useEffect(() => {
    if (preloaded) {
      setSalesData([preloaded]);
      setMemos({ [preloaded.sale.id]: preloaded.sale.memo || '' });
      return;
    }

    if (saleIds.length === 0) return;

    (async () => {
      setLoading(true);
      const supabase = createClient();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const results: SaleData[] = [];

      await Promise.all(saleIds.map(async (id) => {
        const [saleRes, itemsRes, serialsRes] = await Promise.all([
          db.from('offline_sales').select('*').eq('id', id).single(),
          db.from('offline_sale_items').select('*').eq('sale_id', id),
          db.from('product_serials').select('id, serial_number, product_id, sale_item_id').eq('offline_sale_id', id),
        ]);
        if (saleRes.data) {
          results.push({
            sale: saleRes.data,
            items: itemsRes.data || [],
            serials: serialsRes.data || [],
          });
        }
      }));

      // 판매일 순으로 정렬
      results.sort((a, b) => a.sale.sale_number.localeCompare(b.sale.sale_number));
      setSalesData(results);

      const initMemos: Record<string, string> = {};
      results.forEach((r) => { initMemos[r.sale.id] = r.sale.memo || ''; });
      setMemos(initMemos);
      setLoading(false);
    })();
  }, [saleIds, preloaded]);

  // 시리얼 매칭 (sale-detail-panel과 동일 로직)
  function getItemSerials(saleData: SaleData, item: SaleData['items'][0]): string[] {
    return saleData.serials
      .filter((s) =>
        (s.sale_item_id && s.sale_item_id === item.id) ||
        (!s.sale_item_id && s.product_id && s.product_id === item.product_id)
      )
      .map((s) => s.serial_number);
  }

  // 인쇄
  const handlePrint = () => {
    const el = printRef.current;
    if (!el) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>출고 준비표</title>
      <style>
        @page { margin: 15mm; }
        body { font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif; font-size: 12px; color: #000; }
        h1 { text-align: center; font-size: 18px; margin-bottom: 4px; }
        .date { text-align: center; font-size: 11px; color: #888; margin-bottom: 16px; }
        table { width: 100%; border-collapse: collapse; }
        th { background: #f5f5f5; font-weight: 600; font-size: 11px; padding: 6px 8px; border: 1px solid #ddd; }
        td { padding: 5px 8px; border: 1px solid #eee; font-size: 11px; vertical-align: top; }
        .num { text-align: center; font-weight: 600; }
        .qty { text-align: center; }
        .customer { font-weight: 600; }
        .serial { color: #888; font-size: 10px; }
        .memo-cell { min-width: 100px; white-space: pre-wrap; font-size: 10px; color: #555; }
        .group-border td { border-top: 2px solid #333; }
        .footer { text-align: center; margin-top: 16px; font-size: 10px; color: #ccc; }
      </style></head><body>${el.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  const today = new Date().toISOString().slice(0, 10);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '700px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 상단 버튼 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">
            출고 준비표 {salesData.length > 0 && `(${salesData.length}건)`}
          </h3>
          <div className="flex items-center gap-2">
            <button
              onClick={handlePrint}
              disabled={loading || salesData.length === 0}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition disabled:opacity-50"
            >
              <Printer size={12} />
              인쇄
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg leading-none">×</button>
          </div>
        </div>

        {/* 콘텐츠 */}
        <div className="overflow-y-auto flex-1 p-5">
          {loading ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">데이터 로딩 중...</div>
          ) : salesData.length === 0 ? (
            <div className="flex items-center justify-center h-32 text-sm text-neutral-400">선택된 건이 없습니다</div>
          ) : (
            <>
              {/* 메모 수정 영역 (인쇄에는 포함 안 됨) */}
              <div className="mb-4 space-y-2">
                <p className="text-xs text-neutral-500 font-semibold">메모 수정 (인쇄에 반영)</p>
                {salesData.map((sd) => (
                  <div key={sd.sale.id} className="flex items-start gap-2">
                    <span className="text-xs text-neutral-600 w-16 shrink-0 pt-1.5">{sd.sale.customer_name}</span>
                    <input
                      value={memos[sd.sale.id] || ''}
                      onChange={(e) => setMemos({ ...memos, [sd.sale.id]: e.target.value })}
                      placeholder="각인, 서비스 품목 등"
                      className="flex-1 h-8 px-2 rounded-lg border border-neutral-200 text-xs"
                    />
                  </div>
                ))}
              </div>

              {/* 인쇄 영역 */}
              <div ref={printRef}>
                <h1 style={{ textAlign: 'center', fontSize: '18px', fontWeight: 'bold', marginBottom: '4px' }}>
                  MAMORU 출고 준비표
                </h1>
                <p style={{ textAlign: 'center', fontSize: '11px', color: '#888', marginBottom: '16px' }}>{today}</p>

                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr>
                      <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '30px' }}>#</th>
                      <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '80px' }}>고객명</th>
                      <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd' }}>제품명 (시리얼)</th>
                      <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '40px', textAlign: 'center' }}>수량</th>
                      <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '120px' }}>메모</th>
                    </tr>
                  </thead>
                  <tbody>
                    {salesData.map((sd, idx) => (
                      sd.items.map((item, iIdx) => (
                        <tr key={`${sd.sale.id}-${item.id}`}
                          style={iIdx === 0 ? { borderTop: '2px solid #333' } : undefined}>
                          {iIdx === 0 && (
                            <>
                              <td rowSpan={sd.items.length}
                                style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'center', fontWeight: 600, verticalAlign: 'top', fontSize: '11px' }}>
                                {idx + 1}
                              </td>
                              <td rowSpan={sd.items.length}
                                style={{ padding: '5px 8px', border: '1px solid #eee', fontWeight: 600, verticalAlign: 'top', fontSize: '11px' }}>
                                {sd.sale.customer_name}
                              </td>
                            </>
                          )}
                          <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '11px' }}>
                            {item.product_name}
                            {(() => {
                              const serials = getItemSerials(sd, item);
                              if (serials.length === 0) return null;
                              return <span style={{ color: '#888', fontSize: '10px', marginLeft: '4px' }}>({serials.join(', ')})</span>;
                            })()}
                          </td>
                          <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'center', fontSize: '11px' }}>
                            {item.quantity}
                          </td>
                          {iIdx === 0 && (
                            <td rowSpan={sd.items.length}
                              style={{ padding: '5px 8px', border: '1px solid #eee', verticalAlign: 'top', fontSize: '10px', color: '#555', whiteSpace: 'pre-wrap', minWidth: '100px' }}>
                              {memos[sd.sale.id] || ''}
                            </td>
                          )}
                        </tr>
                      ))
                    ))}
                  </tbody>
                </table>

                <p style={{ textAlign: 'center', marginTop: '16px', fontSize: '10px', color: '#ccc' }}>MAMORU</p>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
