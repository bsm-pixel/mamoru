'use client';

import { useEffect, useState } from 'react';
import { X, Award } from 'lucide-react';
import type { DemoPOApi } from './use-demo-po';

const CATEGORIES = [
  { code: 'BL', name: '블런트' },
  { code: 'BK', name: '블런트(코팅)' },
  { code: 'TN', name: '틴닝' },
  { code: 'LG', name: '장가위' },
  { code: 'DR', name: '드라이' },
];

// 디자인 모니터용 SKU 자동 채번 mock — 카테고리당 6번부터 시작 (운영 데이터 없는 가정)
const MOCK_NEXT_SKU: Record<string, number> = {
  BL: 6,
  BK: 4,
  TN: 8,
  LG: 3,
  DR: 2,
};

function nextSkuMock(category: string): string {
  const n = MOCK_NEXT_SKU[category] || 1;
  return `${category}${String(n).padStart(3, '0')}`;
}

export function PromoteModal({
  api,
  itemId,
  onClose,
}: {
  api: DemoPOApi;
  itemId: string | null;
  onClose: () => void;
}) {
  const item = api.po.items.find((it) => it.id === itemId);

  const [category, setCategory] = useState('BL');
  const [sku, setSku] = useState('BL006');
  const [name, setName] = useState('');
  const [price, setPrice] = useState(0);

  useEffect(() => {
    if (item) {
      setCategory('BL');
      const initial = nextSkuMock('BL');
      setSku(initial);
      setName(item.product_name);
      // 매입가의 약 4배 추정 마진 (1688 데모 가이드)
      const krwCost = Math.round(item.unit_price * api.po.exchange_rate);
      setPrice(Math.round((krwCost * 4) / 1000) * 1000);
    }
  }, [item, api.po.exchange_rate]);

  if (!item) return null;

  const handleCategoryChange = (c: string) => {
    setCategory(c);
    setSku(nextSkuMock(c));
  };

  const handleSubmit = () => {
    api.promote(item.id, { sku, name });
    onClose();
  };

  const krwCost = Math.round(item.unit_price * api.po.exchange_rate);
  const marginRate = price > 0 ? Math.round(((price - krwCost) / price) * 100) : 0;

  return (
    <div className="fixed inset-0 z-50 bg-black/40 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-white border-b border-stone-200 px-5 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Award className="text-stone-900" size={18} />
            <span className="font-bold text-stone-900">정식 SKU 승격</span>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="text-stone-400 hover:text-stone-700"
          >
            <X size={18} />
          </button>
        </div>

        <div className="p-5 space-y-4">
          {/* 사입 원본 정보 */}
          <div className="rounded-lg bg-stone-50 border border-stone-200 p-3">
            <div className="text-[10px] text-stone-500 uppercase tracking-wider mb-1">
              사입 원본
            </div>
            <div className="text-xs text-stone-700">
              {item.sticker_no} · ¥{item.unit_price} × {item.quantity} · 매입가 ₩
              {krwCost.toLocaleString()}/개
            </div>
          </div>

          {/* 카테고리 */}
          <Field label="카테고리">
            <div className="grid grid-cols-5 gap-1.5">
              {CATEGORIES.map((c) => (
                <button
                  key={c.code}
                  type="button"
                  onClick={() => handleCategoryChange(c.code)}
                  className={`px-2 py-2 rounded-lg text-[11px] font-bold border ${
                    category === c.code
                      ? 'bg-stone-900 text-white border-stone-900'
                      : 'bg-white text-stone-600 border-stone-200 hover:bg-stone-50'
                  }`}
                >
                  <div>{c.code}</div>
                  <div className="text-[9px] font-normal opacity-70 mt-0.5">{c.name}</div>
                </button>
              ))}
            </div>
          </Field>

          {/* SKU */}
          <Field label="SKU (자동 채번, 수정 가능)">
            <input
              type="text"
              value={sku}
              onChange={(e) => setSku(e.target.value.toUpperCase())}
              className="w-full px-3 py-2 text-sm font-mono font-bold rounded-lg border border-stone-200 bg-white focus:outline-none focus:ring-2 focus:ring-stone-900/10 focus:border-stone-400"
            />
          </Field>

          {/* 정식 제품명 */}
          <Field label="정식 제품명">
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              placeholder="예) MAMORU MB-60 (6.0인치 블런트)"
              className="w-full px-3 py-2 text-sm rounded-lg border border-stone-200 bg-white focus:outline-none focus:ring-2 focus:ring-stone-900/10 focus:border-stone-400"
            />
          </Field>

          {/* 정상 단가 */}
          <Field label="정상 판매 단가 (KRW)">
            <input
              type="number"
              value={price || ''}
              onChange={(e) => setPrice(Number(e.target.value) || 0)}
              className="w-full px-3 py-2 text-sm rounded-lg border border-stone-200 bg-white focus:outline-none focus:ring-2 focus:ring-stone-900/10 focus:border-stone-400"
            />
            <div className="mt-1.5 text-[11px] text-stone-500 flex items-center justify-between">
              <span>매입가 ₩{krwCost.toLocaleString()}</span>
              <span
                className={`font-bold ${
                  marginRate >= 60
                    ? 'text-emerald-700'
                    : marginRate >= 40
                      ? 'text-stone-700'
                      : 'text-rose-600'
                }`}
              >
                마진율 {marginRate}%
              </span>
            </div>
          </Field>

          {/* 사진 복사 옵션 */}
          <label className="flex items-start gap-2 text-xs text-stone-600 select-none cursor-pointer">
            <input type="checkbox" defaultChecked className="mt-0.5" />
            <span>
              1688 입고 사진 {item.inbound_photos.length}장을 정식 제품 이미지로 복사
              <span className="block text-[10px] text-stone-400 mt-0.5">
                (데모 모드: 실제 복사는 일어나지 않습니다)
              </span>
            </span>
          </label>
        </div>

        <div className="sticky bottom-0 bg-white border-t border-stone-200 px-5 py-3 flex items-center justify-end gap-2">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded-lg bg-white border border-stone-200 text-stone-700 text-sm font-medium hover:bg-stone-50"
          >
            취소
          </button>
          <button
            type="button"
            onClick={handleSubmit}
            disabled={!sku || !name}
            className="px-4 py-2 rounded-lg bg-stone-900 text-white text-sm font-bold hover:bg-stone-800 disabled:bg-stone-300 disabled:cursor-not-allowed"
          >
            정식 채택
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <span className="block text-[11px] text-stone-500 mb-1.5 font-medium">{label}</span>
      {children}
    </label>
  );
}
