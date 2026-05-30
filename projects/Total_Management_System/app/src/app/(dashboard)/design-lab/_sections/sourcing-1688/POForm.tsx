'use client';

import { ExternalLink, Plus, Trash2, Sparkles, RotateCcw } from 'lucide-react';
import type { DemoPOApi } from './use-demo-po';

export function POForm({ api }: { api: DemoPOApi }) {
  const { po, totals, updatePO, updateItem, addItem, removeItem, loadSample, reset } = api;

  return (
    <div className="space-y-4">
      {/* 헤더 액션 */}
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <div className="text-xs text-stone-500 mb-0.5">PO 번호</div>
          <div className="font-mono text-sm font-bold text-stone-900">{po.po_number}</div>
        </div>
        <div className="flex items-center gap-2">
          <button
            type="button"
            onClick={loadSample}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-stone-900 text-white text-xs font-medium hover:bg-stone-800"
          >
            <Sparkles size={13} /> 예시 채우기
          </button>
          <button
            type="button"
            onClick={reset}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-white border border-stone-200 text-stone-700 text-xs font-medium hover:bg-stone-50"
          >
            <RotateCcw size={13} /> 초기화
          </button>
        </div>
      </div>

      {/* 매입처 카드 */}
      <div className="rounded-xl border border-stone-200 bg-white p-4">
        <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider mb-3">
          매입처 (1688)
        </div>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
          <Field label="회사명 (중문 가능)">
            <input
              type="text"
              value={po.supplier_name}
              onChange={(e) => updatePO({ supplier_name: e.target.value })}
              placeholder="예) 光达美容工具"
              className={inputCls}
            />
          </Field>
          <Field label="회사 홈 URL">
            <input
              type="url"
              value={po.supplier_url}
              onChange={(e) => updatePO({ supplier_url: e.target.value })}
              placeholder="https://shop0000000.1688.com"
              className={inputCls}
            />
          </Field>
          <Field label="발주일">
            <input
              type="date"
              value={po.order_date}
              onChange={(e) => updatePO({ order_date: e.target.value })}
              className={inputCls}
            />
          </Field>
          <Field label="환율 (CNY → KRW)">
            <input
              type="number"
              value={po.exchange_rate || ''}
              onChange={(e) => updatePO({ exchange_rate: Number(e.target.value) || 0 })}
              placeholder="195"
              className={inputCls}
            />
          </Field>
        </div>
      </div>

      {/* 품목 리스트 */}
      <div className="rounded-xl border border-stone-200 bg-white p-4">
        <div className="flex items-center justify-between mb-3">
          <div className="text-[11px] font-bold text-stone-500 uppercase tracking-wider">
            품목 ({po.items.length}건)
          </div>
          <button
            type="button"
            onClick={addItem}
            className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-stone-100 text-stone-700 text-xs font-medium hover:bg-stone-200"
          >
            <Plus size={12} /> 품목 추가
          </button>
        </div>

        <div className="space-y-3">
          {po.items.map((it, idx) => (
            <div
              key={it.id}
              className="rounded-lg border border-stone-200 bg-stone-50/50 p-3"
            >
              <div className="flex items-center justify-between mb-2">
                <div className="flex items-center gap-2">
                  <span className="inline-flex items-center justify-center w-7 h-7 rounded-full bg-stone-900 text-white text-xs font-bold">
                    {String(idx + 1).padStart(2, '0')}
                  </span>
                  <span className="font-mono text-[11px] text-stone-500">
                    {it.sticker_no}
                  </span>
                </div>
                <button
                  type="button"
                  onClick={() => removeItem(it.id)}
                  disabled={po.items.length === 1}
                  className="text-stone-400 hover:text-rose-500 disabled:opacity-30 disabled:hover:text-stone-400"
                >
                  <Trash2 size={14} />
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-12 gap-2">
                <div className="md:col-span-12">
                  <Field label="1688 상품 URL">
                    <div className="relative">
                      <input
                        type="url"
                        value={it.vendor_url}
                        onChange={(e) => updateItem(it.id, { vendor_url: e.target.value })}
                        placeholder="https://detail.1688.com/offer/..."
                        className={inputCls + ' pr-8'}
                      />
                      {it.vendor_url && (
                        <a
                          href={it.vendor_url}
                          target="_blank"
                          rel="noreferrer"
                          className="absolute right-2 top-1/2 -translate-y-1/2 text-stone-400 hover:text-stone-700"
                        >
                          <ExternalLink size={13} />
                        </a>
                      )}
                    </div>
                  </Field>
                </div>
                <div className="md:col-span-6">
                  <Field label="품목명">
                    <input
                      type="text"
                      value={it.product_name}
                      onChange={(e) => updateItem(it.id, { product_name: e.target.value })}
                      placeholder="예) 6.0인치 일자 가위 (SUS440C)"
                      className={inputCls}
                    />
                  </Field>
                </div>
                <div className="md:col-span-6">
                  <Field label="특징 메모">
                    <input
                      type="text"
                      value={it.features_memo}
                      onChange={(e) => updateItem(it.id, { features_memo: e.target.value })}
                      placeholder="예) 날 광택 양호, 풀너트"
                      className={inputCls}
                    />
                  </Field>
                </div>
                <div className="md:col-span-3">
                  <Field label="MOQ">
                    <input
                      type="number"
                      value={it.moq ?? ''}
                      onChange={(e) =>
                        updateItem(it.id, {
                          moq: e.target.value ? Number(e.target.value) : null,
                        })
                      }
                      placeholder="-"
                      className={inputCls}
                    />
                  </Field>
                </div>
                <div className="md:col-span-3">
                  <Field label="단가 (¥)">
                    <input
                      type="number"
                      value={it.unit_price || ''}
                      onChange={(e) =>
                        updateItem(it.id, { unit_price: Number(e.target.value) || 0 })
                      }
                      className={inputCls}
                    />
                  </Field>
                </div>
                <div className="md:col-span-3">
                  <Field label="수량">
                    <input
                      type="number"
                      value={it.quantity || ''}
                      onChange={(e) =>
                        updateItem(it.id, { quantity: Number(e.target.value) || 0 })
                      }
                      className={inputCls}
                    />
                  </Field>
                </div>
                <div className="md:col-span-3">
                  <Field label="소계">
                    <div className="px-3 py-2 text-sm font-bold text-stone-900 bg-stone-100 rounded-lg">
                      ¥{(it.unit_price * it.quantity).toLocaleString()}
                    </div>
                  </Field>
                </div>
              </div>
            </div>
          ))}
        </div>

        <div className="mt-4 pt-3 border-t border-stone-200 flex items-center justify-between flex-wrap gap-2">
          <div className="text-xs text-stone-500">
            품목 {po.items.length}건 · 총수량 {po.items.reduce((s, it) => s + it.quantity, 0)}개
          </div>
          <div className="text-right">
            <div className="text-xs text-stone-500">합계</div>
            <div className="text-base font-bold text-stone-900">
              ¥{totals.totalCny.toLocaleString()}{' '}
              <span className="text-xs text-stone-500 font-normal">
                ≈ ₩{totals.totalKrw.toLocaleString()}
              </span>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

const inputCls =
  'w-full px-3 py-2 text-sm rounded-lg border border-stone-200 bg-white focus:outline-none focus:ring-2 focus:ring-stone-900/10 focus:border-stone-400';

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <span className="block text-[11px] text-stone-500 mb-1 font-medium">{label}</span>
      {children}
    </label>
  );
}
