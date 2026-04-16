'use client';

import { useRef, useState } from 'react';
import { Printer } from 'lucide-react';
import { useSupplierCatalog } from '@/hooks/use-purchasing';
import { formatKRW } from '@/lib/utils/format';
import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';

interface Props {
  supplierId: string;
  supplierName: string;
  onClose: () => void;
}

type ColumnKey = 'category' | 'sku' | 'product_name' | 'order_name' | 'features' | 'price';

const COLUMN_DEFS: { key: ColumnKey; label: string; default: boolean }[] = [
  { key: 'category', label: '카테고리', default: true },
  { key: 'sku', label: 'SKU', default: true },
  { key: 'product_name', label: '제품명', default: true },
  { key: 'order_name', label: '주문명', default: true },
  { key: 'features', label: '특징', default: true },
  { key: 'price', label: '매입가', default: true },
];

export function CatalogPrintModal({ supplierId, supplierName, onClose }: Props) {
  const printRef = useRef<HTMLDivElement>(null);
  const { data, isLoading } = useSupplierCatalog(supplierId);
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const catalog = data?.catalog || [];

  const [visibleCols, setVisibleCols] = useState<Set<ColumnKey>>(
    new Set(COLUMN_DEFS.filter(c => c.default).map(c => c.key))
  );

  const toggleCol = (key: ColumnKey) => {
    setVisibleCols(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };

  const handlePrint = () => {
    const el = printRef.current;
    if (!el) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>매입품목 카탈로그 - ${supplierName}</title>
      <style>
        @page { margin: 12mm; }
        body { font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif; font-size: 14px; color: #000; }
        h1 { text-align: center; font-size: 22px; font-weight: 800; margin-bottom: 4px; }
        .subtitle { text-align: center; font-size: 13px; color: #888; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
        th { background: #f5f5f5; font-weight: 600; font-size: 13px; padding: 8px 10px; border: 1px solid #ddd; }
        td { padding: 7px 10px; border: 1px solid #eee; font-size: 13px; }
        .footer { text-align: center; margin-top: 24px; font-size: 10px; color: #ccc; }
      </style></head><body>${el.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  if (isLoading) {
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50">
        <div className="bg-white rounded-xl p-8 text-sm text-neutral-400">로딩 중...</div>
      </div>
    );
  }

  const thStyle = { background: '#f5f5f5', fontWeight: 600 as const, fontSize: '13px', padding: '8px 10px', border: '1px solid #ddd' };
  const tdStyle = { padding: '7px 10px', border: '1px solid #eee', fontSize: '13px' };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '800px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">매입품목 카탈로그</h3>
          <div className="flex items-center gap-2">
            <button onClick={handlePrint}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition">
              <Printer size={12} />인쇄
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg">×</button>
          </div>
        </div>

        {/* 컬럼 선택 토글 */}
        <div className="px-4 py-2 border-b border-neutral-100 flex items-center gap-1.5 flex-wrap">
          <span className="text-xs text-neutral-400 mr-1">표시 항목:</span>
          {COLUMN_DEFS.map(col => (
            <button
              key={col.key}
              onClick={() => toggleCol(col.key)}
              className={`px-2.5 py-1 rounded-full text-xs font-medium transition ${
                visibleCols.has(col.key)
                  ? 'bg-neutral-900 text-white'
                  : 'bg-neutral-100 text-neutral-400'
              }`}
            >
              {col.label}
            </button>
          ))}
        </div>

        <div className="overflow-y-auto flex-1 p-5">
          <div ref={printRef}>
            <h1>매입품목 카탈로그</h1>
            <p className="subtitle" style={{ textAlign: 'center', fontSize: '13px', color: '#888', marginBottom: '20px' }}>
              {supplierName} · {catalog.length}개 품목
            </p>

            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ ...thStyle, width: '30px', textAlign: 'center' }}>No</th>
                  {visibleCols.has('category') && <th style={thStyle}>카테고리</th>}
                  {visibleCols.has('sku') && <th style={thStyle}>SKU</th>}
                  {visibleCols.has('product_name') && <th style={thStyle}>제품명</th>}
                  {visibleCols.has('order_name') && <th style={thStyle}>주문명</th>}
                  {visibleCols.has('features') && <th style={thStyle}>특징</th>}
                  {visibleCols.has('price') && <th style={{ ...thStyle, textAlign: 'right' }}>매입가</th>}
                </tr>
              </thead>
              <tbody>
                {catalog.map((entry, idx) => (
                  <tr key={entry.id}>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>{idx + 1}</td>
                    {visibleCols.has('category') && (
                      <td style={tdStyle}>{catLabels[entry.category] || entry.category}</td>
                    )}
                    {visibleCols.has('sku') && (
                      <td style={{ ...tdStyle, fontFamily: 'monospace' }}>{entry.sku}</td>
                    )}
                    {visibleCols.has('product_name') && (
                      <td style={tdStyle}>{entry.product_name}</td>
                    )}
                    {visibleCols.has('order_name') && (
                      <td style={tdStyle}>{entry.order_name || entry.product_name}</td>
                    )}
                    {visibleCols.has('features') && (
                      <td style={{ ...tdStyle, color: '#555' }}>{entry.features || ''}</td>
                    )}
                    {visibleCols.has('price') && (
                      <td style={{ ...tdStyle, textAlign: 'right' }}>{entry.price_purchase > 0 ? formatKRW(entry.price_purchase) : '-'}</td>
                    )}
                  </tr>
                ))}
              </tbody>
            </table>

            <p style={{ textAlign: 'center', marginTop: '24px', fontSize: '10px', color: '#ccc' }}>MAMORU</p>
          </div>
        </div>
      </div>
    </div>
  );
}
