'use client';

import { useRef } from 'react';
import { Printer } from 'lucide-react';
import { useSupplierCatalog, type CatalogEntry } from '@/hooks/use-purchasing';
import { formatKRW } from '@/lib/utils/format';
import { useSetting } from '@/hooks/use-settings';
import { DEFAULT_CAT_LABELS } from '@/lib/utils/setting-defaults';

interface Props {
  supplierId: string;
  supplierName: string;
  onClose: () => void;
}

export function CatalogPrintModal({ supplierId, supplierName, onClose }: Props) {
  const printRef = useRef<HTMLDivElement>(null);
  const { data, isLoading } = useSupplierCatalog(supplierId);
  const catLabels = useSetting<Record<string, string>>('inventory.category_labels', DEFAULT_CAT_LABELS);
  const catalog = data?.catalog || [];

  const handlePrint = () => {
    const el = printRef.current;
    if (!el) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>매입품목 카탈로그 - ${supplierName}</title>
      <style>
        @page { margin: 15mm; }
        body { font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif; font-size: 12px; color: #000; }
        h1 { text-align: center; font-size: 20px; font-weight: 800; margin-bottom: 4px; }
        .subtitle { text-align: center; font-size: 12px; color: #888; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
        th { background: #f5f5f5; font-weight: 600; font-size: 11px; padding: 6px 8px; border: 1px solid #ddd; }
        td { padding: 5px 8px; border: 1px solid #eee; font-size: 11px; }
        .right { text-align: right; }
        .center { text-align: center; }
        .features { font-size: 10px; color: #666; }
        .footer { text-align: center; margin-top: 24px; font-size: 10px; color: #ccc; }
        .total { text-align: right; font-size: 12px; color: #888; margin-bottom: 12px; }
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

  // 카테고리별 그룹핑
  const grouped = new Map<string, CatalogEntry[]>();
  for (const entry of catalog) {
    const cat = entry.category || 'ETC';
    if (!grouped.has(cat)) grouped.set(cat, []);
    grouped.get(cat)!.push(entry);
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '750px', maxHeight: '90vh' }}
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

        <div className="overflow-y-auto flex-1 p-5">
          <div ref={printRef}>
            <h1>매입품목 카탈로그</h1>
            <p className="subtitle">{supplierName} · {catalog.length}개 품목</p>

            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '30px' }}>No</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '70px' }}>카테고리</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '70px' }}>SKU</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd' }}>제품명</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd' }}>주문명</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd' }}>특징</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '80px', textAlign: 'right' }}>매입가</th>
                </tr>
              </thead>
              <tbody>
                {catalog.map((entry, idx) => (
                  <tr key={entry.id}>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'center', fontSize: '11px' }}>{idx + 1}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '10px', textAlign: 'center' }}>{catLabels[entry.category] || entry.category}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '10px', fontFamily: 'monospace' }}>{entry.sku}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '11px' }}>{entry.product_name}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '11px' }}>{entry.order_name || ''}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '10px', color: '#666' }}>{entry.features || ''}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'right', fontSize: '11px' }}>{entry.price_purchase > 0 ? formatKRW(entry.price_purchase) : '-'}</td>
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
