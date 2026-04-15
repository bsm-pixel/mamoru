'use client';

import { useRef } from 'react';
import { Printer, X } from 'lucide-react';
import { usePurchaseOrder } from '@/hooks/use-purchasing';
import { formatKRW } from '@/lib/utils/format';

interface Props {
  purchaseId: string;
  onClose: () => void;
}

export function POPrintModal({ purchaseId, onClose }: Props) {
  const printRef = useRef<HTMLDivElement>(null);
  const { data, isLoading } = usePurchaseOrder(purchaseId);

  const handlePrint = () => {
    const el = printRef.current;
    if (!el) return;
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    printWindow.document.write(`
      <html><head><title>발주서</title>
      <style>
        @page { margin: 15mm; }
        body { font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', sans-serif; font-size: 12px; color: #000; }
        h1 { text-align: center; font-size: 22px; font-weight: 800; margin-bottom: 4px; letter-spacing: 8px; }
        .subtitle { text-align: center; font-size: 11px; color: #888; margin-bottom: 24px; }
        .info-grid { display: flex; justify-content: space-between; margin-bottom: 20px; }
        .info-box { width: 48%; }
        .info-box h3 { font-size: 11px; font-weight: 700; color: #666; margin-bottom: 8px; border-bottom: 1px solid #ddd; padding-bottom: 4px; }
        .info-row { display: flex; font-size: 11px; padding: 3px 0; }
        .info-label { width: 60px; color: #888; flex-shrink: 0; }
        .info-value { flex: 1; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 16px; }
        th { background: #f5f5f5; font-weight: 600; font-size: 11px; padding: 6px 8px; border: 1px solid #ddd; }
        td { padding: 5px 8px; border: 1px solid #eee; font-size: 11px; }
        .num { text-align: center; }
        .right { text-align: right; }
        .total-section { margin-top: 8px; }
        .total-row { display: flex; justify-content: space-between; padding: 3px 0; font-size: 12px; }
        .total-row.bold { font-weight: 700; font-size: 13px; border-top: 2px solid #333; padding-top: 6px; margin-top: 4px; }
        .memo-section { margin-top: 16px; padding: 12px; background: #f9f9f9; border-radius: 4px; font-size: 11px; color: #555; }
        .footer { text-align: center; margin-top: 24px; font-size: 10px; color: #ccc; }
      </style></head><body>${el.innerHTML}</body></html>
    `);
    printWindow.document.close();
    printWindow.print();
  };

  if (isLoading || !data) {
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50">
        <div className="bg-white rounded-xl p-8 text-sm text-neutral-400">로딩 중...</div>
      </div>
    );
  }

  const { order: po, items } = data;
  const vatType = ((po as Record<string, unknown>).vat_type as string) || 'included';

  // 부가세 계산
  const itemTotal = items.reduce((s, i) => s + i.total_price, 0);
  let supply = 0, vat = 0, payment = 0;
  if (vatType === 'separate') {
    supply = itemTotal;
    vat = Math.round(itemTotal * 0.1);
    payment = itemTotal + vat;
  } else if (vatType === 'none') {
    supply = itemTotal;
    vat = 0;
    payment = itemTotal;
  } else {
    supply = Math.round(itemTotal / 1.1);
    vat = itemTotal - supply;
    payment = itemTotal;
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: '700px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">발주서 미리보기</h3>
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
            <h1>발 주 서</h1>
            <p className="subtitle">{po.po_number} · {po.order_date}</p>

            <div className="info-grid" style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '20px' }}>
              <div className="info-box" style={{ width: '48%' }}>
                <h3 style={{ fontSize: '11px', fontWeight: 700, color: '#666', marginBottom: '8px', borderBottom: '1px solid #ddd', paddingBottom: '4px' }}>매입처 정보</h3>
                <div style={{ fontSize: '11px', lineHeight: '1.8' }}>
                  <div><span style={{ color: '#888', display: 'inline-block', width: '50px' }}>업체명</span> {po.supplier_name} <span style={{ color: '#888' }}>귀중</span></div>
                  {po.expected_date && <div><span style={{ color: '#888', display: 'inline-block', width: '50px' }}>납기일</span> {po.expected_date}</div>}
                </div>
              </div>
              <div className="info-box" style={{ width: '48%' }}>
                <h3 style={{ fontSize: '11px', fontWeight: 700, color: '#666', marginBottom: '8px', borderBottom: '1px solid #ddd', paddingBottom: '4px' }}>발주 정보</h3>
                <div style={{ fontSize: '11px', lineHeight: '1.8' }}>
                  <div><span style={{ color: '#888', display: 'inline-block', width: '50px' }}>발주일</span> {po.order_date}</div>
                  <div><span style={{ color: '#888', display: 'inline-block', width: '50px' }}>부가세</span> {vatType === 'included' ? '포함' : vatType === 'separate' ? '별도' : '미적용'}</div>
                </div>
              </div>
            </div>

            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: '16px' }}>
              <thead>
                <tr>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '30px' }}>No</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd' }}>주문품목</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '80px', textAlign: 'right' }}>단가</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '50px', textAlign: 'center' }}>수량</th>
                  <th style={{ background: '#f5f5f5', fontWeight: 600, fontSize: '11px', padding: '6px 8px', border: '1px solid #ddd', width: '90px', textAlign: 'right' }}>금액</th>
                </tr>
              </thead>
              <tbody>
                {items.map((item, idx) => (
                  <tr key={item.id}>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'center', fontSize: '11px' }}>{idx + 1}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', fontSize: '11px' }}>{item.product_name}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'right', fontSize: '11px' }}>{formatKRW(item.unit_price)}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'center', fontSize: '11px' }}>{item.quantity}</td>
                    <td style={{ padding: '5px 8px', border: '1px solid #eee', textAlign: 'right', fontSize: '11px' }}>{formatKRW(item.total_price)}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            {/* 합계 */}
            <div style={{ marginTop: '8px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0', fontSize: '12px' }}>
                <span>공급가액</span><span>{formatKRW(supply)}</span>
              </div>
              {vatType !== 'none' && (
                <div style={{ display: 'flex', justifyContent: 'space-between', padding: '3px 0', fontSize: '12px' }}>
                  <span>부가세 (10%)</span><span>{formatKRW(vat)}</span>
                </div>
              )}
              <div style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0 0', fontSize: '14px', fontWeight: 700, borderTop: '2px solid #333', marginTop: '4px' }}>
                <span>합계</span><span>{formatKRW(payment)}</span>
              </div>
            </div>

            {po.memo && (
              <div style={{ marginTop: '16px', padding: '12px', background: '#f9f9f9', borderRadius: '4px', fontSize: '11px', color: '#555' }}>
                <strong>메모:</strong> {po.memo}
              </div>
            )}

            <p style={{ textAlign: 'center', marginTop: '24px', fontSize: '10px', color: '#ccc' }}>MAMORU</p>
          </div>
        </div>
      </div>
    </div>
  );
}
