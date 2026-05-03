'use client';

/**
 * 재고조사 인쇄 모달
 *
 * 사장님이 종이를 들고 다니며 실재고 카운트할 때 사용.
 * - 화면에서 적용한 카테고리/필터/정렬 그대로 인쇄에 반영
 * - 카테고리별 자동 그룹화 + 그룹별 소계
 * - 실측 빈 칸 + 비고란 (체크용)
 * - window.open() 새 탭 패턴 (po-print-modal와 동일)
 */

import { useRef, useMemo } from 'react';
import { Printer } from 'lucide-react';
import type { InventoryItem } from '@/hooks/use-inventory';

interface Props {
  items: InventoryItem[];
  categoryLabel: string;       // 화면 카테고리 칩 라벨 (예: "가위", "전체")
  categoryLabels: Record<string, string>;  // 코드 → 라벨 매핑 (BL → "블런트")
  filters: {
    search?: string;
    lowStockOnly: boolean;
    hideUnused: boolean;
    sortKey: string;
    sortAsc: boolean;
  };
  onClose: () => void;
}

const SORT_LABEL: Record<string, string> = {
  name: '이름순',
  stock_quantity: '재고순',
  pending_quantity: '미입고순',
  value: '원가순',
};

export function InventoryPrintModal({ items, categoryLabel, categoryLabels, filters, onClose }: Props) {
  const printRef = useRef<HTMLDivElement>(null);

  // 카테고리별 자동 그룹화 (현재 정렬 순서 유지)
  const grouped = useMemo(() => {
    const map = new Map<string, InventoryItem[]>();
    for (const item of items) {
      const cat = item.category || '기타';
      if (!map.has(cat)) map.set(cat, []);
      map.get(cat)!.push(item);
    }
    return Array.from(map.entries()); // [[category, items], ...]
  }, [items]);

  const today = new Date();
  const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
  const totalCount = items.reduce((s, i) => s + (i.stock_quantity > 0 ? i.stock_quantity : 0), 0);

  // 필터 요약 표시
  const filterSummary: string[] = [];
  filterSummary.push(`카테고리: ${categoryLabel}`);
  filterSummary.push(`정렬: ${SORT_LABEL[filters.sortKey] || filters.sortKey} ${filters.sortAsc ? '↑' : '↓'}`);
  if (filters.lowStockOnly) filterSummary.push('저재고만');
  if (filters.search) filterSummary.push(`검색: "${filters.search}"`);

  const handlePrint = () => {
    const el = printRef.current;
    if (!el) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>재고조사 — ${dateStr}</title>
      <style>
        @page { margin: 12mm; }
        body { font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif; font-size: 11px; color: #000; }
        h1 { text-align: center; font-size: 20px; font-weight: 800; margin: 0 0 4px; letter-spacing: 4px; }
        .subtitle { text-align: center; font-size: 10px; color: #666; margin-bottom: 16px; }
        .filter-info { font-size: 10px; color: #555; margin-bottom: 12px; padding: 6px 10px; background: #f5f5f5; border-radius: 3px; }
        .group-header { background: #1a1a1a; color: #fff; padding: 6px 10px; font-size: 12px; font-weight: 700; margin-top: 16px; border-radius: 3px; }
        .group-subtotal { font-size: 10px; color: #aaa; font-weight: 400; }
        table { width: 100%; border-collapse: collapse; margin-top: 4px; }
        th { background: #f5f5f5; font-weight: 600; font-size: 10px; padding: 5px 6px; border: 1px solid #ddd; text-align: center; }
        td { padding: 5px 6px; border: 1px solid #eee; font-size: 11px; }
        td.num { text-align: center; }
        td.name { font-weight: 600; }
        td.sku { color: #666; font-size: 10px; }
        td.measure { background: #fffacd; min-width: 50px; }
        .total-row { background: #f9f9f9; font-weight: 700; }
        .footer { margin-top: 20px; padding-top: 12px; border-top: 2px solid #333; font-size: 11px; }
        .memo-area { margin-top: 12px; padding: 12px; border: 1px dashed #999; font-size: 10px; color: #666; min-height: 50px; }
      </style></head><body>${el.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '780px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">재고조사 인쇄 미리보기</h3>
          <div className="flex items-center gap-2">
            <button
              onClick={handlePrint}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition"
            >
              <Printer size={12} />인쇄
            </button>
            <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg">×</button>
          </div>
        </div>

        {/* 미리보기 */}
        <div className="overflow-y-auto flex-1 p-5">
          <div ref={printRef}>
            <h1>MAMORU 재고조사</h1>
            <p className="subtitle" style={{ textAlign: 'center', fontSize: '10px', color: '#666', marginBottom: '16px' }}>
              {dateStr} · 총 {items.length}종 · {totalCount}개
            </p>

            <div className="filter-info" style={{ fontSize: '10px', color: '#555', marginBottom: '12px', padding: '6px 10px', background: '#f5f5f5', borderRadius: '3px' }}>
              <strong>적용 필터:</strong> {filterSummary.join(' · ')}
            </div>

            {grouped.map(([cat, list]) => {
              const groupTotal = list.reduce((s, i) => s + (i.stock_quantity > 0 ? i.stock_quantity : 0), 0);
              return (
                <div key={cat}>
                  <div className="group-header" style={{ background: '#1a1a1a', color: '#fff', padding: '6px 10px', fontSize: '12px', fontWeight: 700, marginTop: '16px', borderRadius: '3px' }}>
                    {categoryLabels[cat] || cat}
                    <span className="group-subtotal" style={{ fontSize: '10px', color: '#aaa', fontWeight: 400, marginLeft: '8px' }}>
                      ({list.length}종 · 소계 {groupTotal}개)
                    </span>
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '4px' }}>
                    <thead>
                      <tr>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '32px' }}>#</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', textAlign: 'left' }}>제품명</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '90px' }}>SKU</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '50px' }}>현재고</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '40px' }}>보관</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '40px' }}>준비</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '40px' }}>디스</th>
                        <th style={{ background: '#fffacd', fontWeight: 700, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '60px' }}>실측</th>
                        <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '10px', padding: '5px 6px', border: '1px solid #ddd', width: '80px' }}>비고</th>
                      </tr>
                    </thead>
                    <tbody>
                      {list.map((item, idx) => {
                        const sku = (item.sku || '').startsWith('IW-') ? '' : (item.sku || '');
                        return (
                          <tr key={item.id}>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', textAlign: 'center', fontSize: '10px', color: '#888' }}>{idx + 1}</td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', fontSize: '11px', fontWeight: 600 }}>{item.name}</td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', fontSize: '10px', color: '#666' }}>{sku}</td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', textAlign: 'center', fontSize: '11px', fontWeight: 700 }}>
                              {item.stock_quantity === -1 ? '-' : item.stock_quantity}
                            </td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', textAlign: 'center', fontSize: '10px' }}>{item.zone_raw || ''}</td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', textAlign: 'center', fontSize: '10px' }}>{item.zone_ready || ''}</td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', textAlign: 'center', fontSize: '10px' }}>{item.zone_display || ''}</td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee', background: '#fffacd' }}></td>
                            <td style={{ padding: '5px 6px', border: '1px solid #eee' }}></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              );
            })}

            {/* 총계 + 메모 */}
            <div className="footer" style={{ marginTop: '20px', paddingTop: '12px', borderTop: '2px solid #333', fontSize: '11px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between' }}>
                <span><strong>전체 종 수:</strong> {items.length}종</span>
                <span><strong>전체 수량:</strong> {totalCount}개</span>
              </div>
            </div>

            <div className="memo-area" style={{ marginTop: '12px', padding: '12px', border: '1px dashed #999', fontSize: '10px', color: '#666', minHeight: '50px' }}>
              비고:
            </div>

            <p style={{ textAlign: 'center', marginTop: '16px', fontSize: '9px', color: '#ccc' }}>MAMORU · {dateStr}</p>
          </div>
        </div>
      </div>
    </div>
  );
}
