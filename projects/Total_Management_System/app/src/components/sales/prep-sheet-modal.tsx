'use client';

import { useState, useRef, useEffect, useMemo } from 'react';
import { Printer } from 'lucide-react';
import { createClient } from '@/lib/supabase/client';

/** 2026-05-26 Phase D: B2C(sale) + B2B(delivery) 통합 준비표 출력
 *  2026-07-23: 트레이형 추가 — A4 1장에 주문 2건(좌·우), 중앙 절취선. 반쪽(105×297mm) 트레이용. */
interface SaleData {
  sourceType: 'sale' | 'delivery';
  sale: {
    id: string;
    sale_number: string;       // delivery 의 경우 dl_number 매핑
    sale_date: string;          // delivery 의 경우 delivery_date 매핑
    customer_name: string;
    customer_phone?: string | null;
    customer_id?: string | null;
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
  /** 2026-05-26 Phase D: 거래처 납품 ID — B2B 카드 체크 시 합산 표시 */
  deliveryIds?: string[];
  /** 단건 모드: 이미 로드된 데이터 직접 전달 */
  preloaded?: SaleData;
  onClose: () => void;
}

type PrepMode = 'list' | 'tray';
function esc(s: unknown): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function PrepSheetModal({ saleIds, deliveryIds = [], preloaded, onClose }: PrepSheetModalProps) {
  const printRef = useRef<HTMLDivElement>(null);
  const [salesData, setSalesData] = useState<SaleData[]>([]);
  const [memos, setMemos] = useState<Record<string, string>>({});
  const [addresses, setAddresses] = useState<Record<string, string>>({}); // customer_id → 주소 문자열
  const [loading, setLoading] = useState(false);
  const [mode, setMode] = useState<PrepMode>('list'); // 기존 동작 보존: 기본 리스트형

  // 데이터 로딩
  useEffect(() => {
    if (preloaded) {
      setSalesData([preloaded]);
      setMemos({ [preloaded.sale.id]: preloaded.sale.memo || '' });
      // 단건도 주소 로드 (트레이형)
      const cid = preloaded.sale.customer_id;
      if (cid) {
        (async () => {
          const db = createClient() as unknown as { from: (t: string) => { select: (c: string) => { eq: (k: string, v: string) => { single: () => Promise<{ data: { postcode?: string; address_road?: string; address_detail?: string } | null }> } } } };
          const { data } = await db.from('customers').select('id, postcode, address_road, address_detail').eq('id', cid).single();
          if (data) setAddresses({ [cid]: [data.postcode ? `(${data.postcode})` : '', data.address_road, data.address_detail].filter(Boolean).join(' ') });
        })();
      }
      return;
    }

    if (saleIds.length === 0 && deliveryIds.length === 0) return;

    (async () => {
      setLoading(true);
      const supabase = createClient();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const db = supabase as any;
      const results: SaleData[] = [];

      // B2C — offline_sales + items + serials
      await Promise.all(saleIds.map(async (id) => {
        const [saleRes, itemsRes, serialsRes] = await Promise.all([
          db.from('offline_sales').select('*').eq('id', id).single(),
          db.from('offline_sale_items').select('*').eq('sale_id', id),
          db.from('product_serials').select('id, serial_number, product_id, sale_item_id').eq('offline_sale_id', id),
        ]);
        if (saleRes.data) {
          results.push({
            sourceType: 'sale',
            sale: saleRes.data,
            items: itemsRes.data || [],
            serials: serialsRes.data || [],
          });
        }
      }));

      // 2026-05-26 Phase D: B2B — deliveries + delivery_items (시리얼 미부여)
      await Promise.all(deliveryIds.map(async (id) => {
        const [dlRes, dlItemsRes] = await Promise.all([
          db.from('deliveries').select('*').eq('id', id).single(),
          db.from('delivery_items').select('*').eq('delivery_id', id),
        ]);
        if (dlRes.data) {
          const d = dlRes.data;
          results.push({
            sourceType: 'delivery',
            sale: {
              ...d,
              sale_number: d.dl_number || d.id,
              sale_date: d.delivery_date || d.created_at || '',
              customer_name: d.company_name || d.customer_name || '미지정',
              customer_phone: d.customer_phone,
              memo: d.memo,
            },
            items: (dlItemsRes.data || []).map((it: Record<string, unknown>) => ({
              ...it,
              id: it.id as string,
              product_name: (it.product_name as string) || '',
              quantity: (it.quantity as number) || 0,
              unit_price: (it.unit_price as number) || 0,
            })),
            serials: [],
          });
        }
      }));

      // 정렬: B2C(sale_number) → B2B(dl_number) 순
      results.sort((a, b) => {
        if (a.sourceType !== b.sourceType) return a.sourceType === 'sale' ? -1 : 1;
        return a.sale.sale_number.localeCompare(b.sale.sale_number);
      });
      setSalesData(results);

      const initMemos: Record<string, string> = {};
      results.forEach((r) => { initMemos[r.sale.id] = r.sale.memo || ''; });
      setMemos(initMemos);

      // 트레이형 슬립용 — B2C 고객 주소 배치 로드
      const custIds = [...new Set(results.filter((r) => r.sourceType === 'sale' && r.sale.customer_id).map((r) => r.sale.customer_id as string))];
      if (custIds.length > 0) {
        const { data: custs } = await db.from('customers').select('id, postcode, address_road, address_detail').in('id', custIds);
        const addrMap: Record<string, string> = {};
        (custs || []).forEach((c: { id: string; postcode?: string; address_road?: string; address_detail?: string }) => {
          addrMap[c.id] = [c.postcode ? `(${c.postcode})` : '', c.address_road, c.address_detail].filter(Boolean).join(' ');
        });
        setAddresses(addrMap);
      }
      setLoading(false);
    })();
  }, [saleIds, deliveryIds, preloaded]);

  // 시리얼 매칭 (sale-detail-panel과 동일 로직)
  function getItemSerials(saleData: SaleData, item: SaleData['items'][0]): string[] {
    return saleData.serials
      .filter((s) =>
        (s.sale_item_id && s.sale_item_id === item.id) ||
        (!s.sale_item_id && s.product_id && s.product_id === item.product_id)
      )
      .map((s) => s.serial_number);
  }

  const today = new Date().toISOString().slice(0, 10);

  // ── 트레이형 인쇄 HTML: A4 1장 = 슬립 2개(좌·우), 중앙 절취선 ──
  const buildTrayHtml = () => {
    const slip = (sd: SaleData) => {
      const isD = sd.sourceType === 'delivery';
      const addr = sd.sale.customer_id ? (addresses[sd.sale.customer_id] || '') : '';
      const itemsHtml = sd.items.map((it) => {
        const ser = getItemSerials(sd, it);
        return `<div class="it">☐ <b>${esc(it.product_name)}</b>${ser.length ? ` <span class="ser">(${esc(ser.join(', '))})</span>` : ''}<span class="q">×${it.quantity}</span></div>`;
      }).join('') || '<div class="it muted">품목 없음</div>';
      const memo = memos[sd.sale.id] || '';
      return `<div class="slip">
        <div class="hd">MAMORU 출고 준비표</div>
        <div class="ono">${esc(sd.sale.sale_number)}${isD ? ' <span class="tag">거래처</span>' : ''}</div>
        <div class="dt">${esc((sd.sale.sale_date || '').slice(0, 10))}</div>
        <div class="sec">
          <div class="cust">${esc(sd.sale.customer_name)}${isD ? '' : '<span class="nim"> 님</span>'}</div>
          ${sd.sale.customer_phone ? `<div class="ph">${esc(sd.sale.customer_phone)}</div>` : ''}
          ${addr ? `<div class="addr">${esc(addr)}</div>` : ''}
        </div>
        <div class="sec items">${itemsHtml}</div>
        ${memo ? `<div class="sec memo"><span class="lbl">메모</span> ${esc(memo)}</div>` : ''}
        <div class="chk">☐ 포장 완료&nbsp;&nbsp;&nbsp;☐ 발송 완료</div>
        <div class="ft">MAMORU</div>
      </div>`;
    };
    const pages: string[] = [];
    for (let i = 0; i < salesData.length; i += 2) {
      const left = slip(salesData[i]);
      const right = salesData[i + 1] ? slip(salesData[i + 1]) : '<div class="slip empty"></div>';
      pages.push(`<div class="page">${left}${right}</div>`);
    }
    return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>출고 준비표 (트레이)</title>
    <style>
      @page { size: A4; margin: 0; }
      * { box-sizing: border-box; margin: 0; }
      body { font-family:'Noto Sans KR','Apple SD Gothic Neo',sans-serif; color:#000; }
      .page { width:210mm; height:297mm; display:flex; page-break-after:always; }
      .page:last-child { page-break-after:auto; }
      /* 반쪽 = 105mm. 좌측 슬립 오른쪽 테두리 = 중앙 절취선 */
      .slip { width:105mm; height:297mm; padding:11mm 8mm 9mm; display:flex; flex-direction:column;
              border-right:1px dashed #888; overflow:hidden; }
      .slip.empty, .page > .slip:last-child { border-right:0; }
      .hd { font-size:10px; letter-spacing:2px; color:#666; }
      .ono { font-family:'Courier New',monospace; font-weight:800; font-size:17px; margin-top:3mm; letter-spacing:.5px; }
      .tag { font-size:9px; background:#000; color:#fff; padding:1px 5px; border-radius:3px; font-family:sans-serif; vertical-align:middle; }
      .dt { font-size:11px; color:#888; margin-top:1mm; }
      .sec { border-top:1px solid #000; margin-top:4mm; padding-top:3mm; }
      .cust { font-size:23px; font-weight:800; line-height:1.15; }
      .nim { font-size:13px; font-weight:400; color:#666; }
      .ph { font-size:14px; margin-top:1.5mm; }
      .addr { font-size:12px; color:#333; margin-top:1.5mm; line-height:1.45; }
      .items .it { font-size:13px; padding:1.8mm 0; border-bottom:1px dotted #ddd; overflow:hidden; }
      .items .it:last-child { border-bottom:0; }
      .it .ser { color:#888; font-size:10px; }
      .it .q { float:right; font-weight:700; }
      .it.muted { color:#aaa; }
      .memo { font-size:12px; white-space:pre-wrap; line-height:1.5; }
      .memo .lbl { font-weight:700; }
      .chk { margin-top:auto; border-top:2px solid #000; padding-top:3mm; font-size:15px; font-weight:600; }
      .ft { text-align:center; font-size:9px; color:#ccc; margin-top:3mm; letter-spacing:3px; }
      @media print { body { -webkit-print-color-adjust:exact; print-color-adjust:exact; } }
    </style></head><body>${pages.join('')}</body></html>`;
  };

  const trayHtml = useMemo(() => (mode === 'tray' && salesData.length > 0 ? buildTrayHtml() : ''),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [mode, salesData, memos, addresses]);

  // 인쇄
  const handlePrint = () => {
    const printWindow = window.open('', '_blank');
    if (!printWindow) return;
    if (mode === 'tray') {
      printWindow.document.write(trayHtml || buildTrayHtml());
    } else {
      const el = printRef.current;
      if (!el) { printWindow.close(); return; }
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
        </style></head><body>${el.innerHTML}</body></html>
      `);
    }
    printWindow.document.close();
    printWindow.print();
  };

  // 트레이 미리보기 스케일 (A4 210mm ≈ 793px)
  const previewW = 300;
  const scale = previewW / (210 * 3.7795);
  const previewH = Math.min(297 * 3.7795 * scale * Math.ceil(salesData.length / 2), 460);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-2xl flex flex-col"
        style={{ width: mode === 'tray' ? '760px' : '700px', maxHeight: '90vh' }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* 상단: 모드 토글 + 인쇄 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">
            출고 준비표 {salesData.length > 0 && `(${salesData.length}건)`}
          </h3>
          <div className="flex items-center gap-2">
            {/* 리스트형 / 트레이형 */}
            <div className="flex rounded-lg border border-neutral-200 overflow-hidden text-xs">
              <button onClick={() => setMode('list')}
                className={`px-3 py-1.5 font-semibold transition ${mode === 'list' ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500 hover:bg-neutral-50'}`}>리스트형</button>
              <button onClick={() => setMode('tray')}
                className={`px-3 py-1.5 font-semibold transition ${mode === 'tray' ? 'bg-neutral-900 text-white' : 'bg-white text-neutral-500 hover:bg-neutral-50'}`}>트레이형</button>
            </div>
            <button
              onClick={handlePrint}
              disabled={loading || salesData.length === 0}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-md bg-neutral-900 text-xs text-white hover:bg-neutral-800 transition disabled:opacity-50"
            >
              <Printer size={12} /> 인쇄
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
              {/* 메모 수정 (리스트·트레이 공통 반영) */}
              <div className="mb-4 space-y-2">
                <p className="text-xs text-neutral-500 font-semibold">메모 수정 (인쇄에 반영)</p>
                {salesData.map((sd) => (
                  <div key={sd.sale.id} className="flex items-start gap-2">
                    <span className="text-xs text-neutral-600 w-20 shrink-0 pt-1.5">{sd.sale.customer_name} <span className="text-neutral-400">님</span></span>
                    <input
                      value={memos[sd.sale.id] || ''}
                      onChange={(e) => setMemos({ ...memos, [sd.sale.id]: e.target.value })}
                      placeholder="각인, 서비스 품목 등"
                      className="flex-1 h-8 px-2 rounded-lg border border-neutral-200 text-xs"
                    />
                  </div>
                ))}
              </div>

              {mode === 'tray' ? (
                /* ── 트레이형 미리보기 (실제 인쇄물과 동일 HTML) ── */
                <div>
                  <p className="text-xs text-neutral-500 mb-2">
                    A4 1장에 <b>2건</b> · 가운데 <b>점선에서 세로로 절취</b> → 반쪽(105×297mm) 트레이용 · 총 <b>{Math.ceil(salesData.length / 2)}장</b>
                    {salesData.length % 2 === 1 && <span className="text-neutral-400"> (마지막 장은 왼쪽만)</span>}
                  </p>
                  <div className="border border-neutral-200 rounded-lg overflow-hidden bg-neutral-100"
                    style={{ width: previewW, height: previewH, overflowY: 'auto' }}>
                    <div style={{ width: 210 * 3.7795, transform: `scale(${scale})`, transformOrigin: 'top left' }}>
                      <iframe title="트레이 미리보기" srcDoc={trayHtml} scrolling="no"
                        style={{ width: 210 * 3.7795, height: 297 * 3.7795 * Math.ceil(salesData.length / 2), border: 0, display: 'block' }} />
                    </div>
                  </div>
                </div>
              ) : (
                /* ── 리스트형 (기존) ── */
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
                                  {sd.sourceType === 'delivery' ? (
                                    <span style={{ color: '#666', fontWeight: 'normal', fontSize: '10px', marginLeft: '4px' }}>[거래처]</span>
                                  ) : (
                                    <span style={{ color: '#888', fontWeight: 'normal' }}> 님</span>
                                  )}
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
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
