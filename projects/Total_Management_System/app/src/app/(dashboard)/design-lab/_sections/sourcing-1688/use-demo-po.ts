'use client';

import { useCallback, useMemo, useState } from 'react';
import type { DemoPO, DemoPOItem, InspectionStatus } from './types';

const today = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

const todayCompact = () => today().replaceAll('-', '');

const stickerNo = (poNumber: string, idx: number) =>
  `${poNumber}-${String(idx + 1).padStart(3, '0')}`;

const newItem = (poNumber: string, idx: number): DemoPOItem => ({
  id: `item-${Math.floor(Math.random() * 100000)}-${idx}`,
  vendor_url: '',
  product_name: '',
  features_memo: '',
  moq: null,
  unit_price: 0,
  quantity: 1,
  sticker_no: stickerNo(poNumber, idx),
  inbound_photos: [],
  inbound_memo: '',
  inspection_status: 'pending',
});

const emptyPO = (): DemoPO => {
  const po_number = `PO-${todayCompact()}-DEMO`;
  return {
    po_number,
    supplier_name: '',
    supplier_url: '',
    order_date: today(),
    exchange_rate: 195,
    items: [newItem(po_number, 0)],
  };
};

const samplePO = (): DemoPO => {
  const po_number = `PO-${todayCompact()}-DEMO`;
  return {
    po_number,
    supplier_name: '光达美容工具',
    supplier_url: 'https://shop1234567.1688.com',
    order_date: today(),
    exchange_rate: 195,
    items: [
      {
        id: 'item-sample-1',
        vendor_url: 'https://detail.1688.com/offer/700000000001.html',
        product_name: '6.0인치 일자 가위 (SUS440C)',
        features_memo: '날 광택 양호 / 손잡이 무도금 / 풀너트',
        moq: 30,
        unit_price: 78,
        quantity: 30,
        sticker_no: stickerNo(po_number, 0),
        inbound_photos: [],
        inbound_memo: '',
        inspection_status: 'pending',
      },
      {
        id: 'item-sample-2',
        vendor_url: 'https://detail.1688.com/offer/700000000002.html',
        product_name: '5.5인치 틴닝 가위 (35날)',
        features_memo: '커팅율 약 28~32%, 좌수전용 옵션 있음',
        moq: 20,
        unit_price: 96,
        quantity: 20,
        sticker_no: stickerNo(po_number, 1),
        inbound_photos: [],
        inbound_memo: '',
        inspection_status: 'pending',
      },
      {
        id: 'item-sample-3',
        vendor_url: 'https://detail.1688.com/offer/700000000003.html',
        product_name: '7.0인치 장가위 (백조형 손잡이)',
        features_memo: '무게 약 62g, 케이스 별도',
        moq: 10,
        unit_price: 145,
        quantity: 10,
        sticker_no: stickerNo(po_number, 2),
        inbound_photos: [],
        inbound_memo: '',
        inspection_status: 'pending',
      },
    ],
  };
};

export type DemoPOApi = {
  po: DemoPO;
  selectedItemId: string | null;
  select: (itemId: string | null) => void;
  loadSample: () => void;
  reset: () => void;
  updatePO: (patch: Partial<Omit<DemoPO, 'items'>>) => void;
  updateItem: (itemId: string, patch: Partial<DemoPOItem>) => void;
  addItem: () => void;
  removeItem: (itemId: string) => void;
  addPhoto: (itemId: string, dataUrl: string) => void;
  removePhoto: (itemId: string, idx: number) => void;
  completeMatch: (itemId: string, memo: string) => void;
  promote: (
    itemId: string,
    payload: { sku: string; name: string }
  ) => void;
  reject: (itemId: string) => void;
  setStatus: (itemId: string, status: InspectionStatus) => void;
  totals: {
    matched: number;
    promoted: number;
    rejected: number;
    pending: number;
    totalCny: number;
    totalKrw: number;
  };
};

export function useDemoPO(): DemoPOApi {
  const [po, setPO] = useState<DemoPO>(() => emptyPO());
  const [selectedItemId, setSelectedItemId] = useState<string | null>(null);

  const select = useCallback((itemId: string | null) => {
    setSelectedItemId(itemId);
  }, []);

  const rebuildStickerNos = (items: DemoPOItem[], po_number: string): DemoPOItem[] =>
    items.map((it, idx) => ({ ...it, sticker_no: stickerNo(po_number, idx) }));

  const loadSample = useCallback(() => {
    setPO(samplePO());
    setSelectedItemId(null);
  }, []);

  const reset = useCallback(() => {
    setPO(emptyPO());
    setSelectedItemId(null);
  }, []);

  const updatePO = useCallback((patch: Partial<Omit<DemoPO, 'items'>>) => {
    setPO((cur) => ({ ...cur, ...patch }));
  }, []);

  const updateItem = useCallback((itemId: string, patch: Partial<DemoPOItem>) => {
    setPO((cur) => ({
      ...cur,
      items: cur.items.map((it) => (it.id === itemId ? { ...it, ...patch } : it)),
    }));
  }, []);

  const addItem = useCallback(() => {
    setPO((cur) => ({
      ...cur,
      items: [...cur.items, newItem(cur.po_number, cur.items.length)],
    }));
  }, []);

  const removeItem = useCallback((itemId: string) => {
    setPO((cur) => {
      const filtered = cur.items.filter((it) => it.id !== itemId);
      return { ...cur, items: rebuildStickerNos(filtered, cur.po_number) };
    });
  }, []);

  const addPhoto = useCallback((itemId: string, dataUrl: string) => {
    setPO((cur) => ({
      ...cur,
      items: cur.items.map((it) =>
        it.id === itemId
          ? { ...it, inbound_photos: [...it.inbound_photos, dataUrl].slice(0, 5) }
          : it
      ),
    }));
  }, []);

  const removePhoto = useCallback((itemId: string, idx: number) => {
    setPO((cur) => ({
      ...cur,
      items: cur.items.map((it) =>
        it.id === itemId
          ? {
              ...it,
              inbound_photos: it.inbound_photos.filter((_, i) => i !== idx),
            }
          : it
      ),
    }));
  }, []);

  const completeMatch = useCallback((itemId: string, memo: string) => {
    setPO((cur) => ({
      ...cur,
      items: cur.items.map((it) =>
        it.id === itemId
          ? {
              ...it,
              inbound_memo: memo,
              inspection_status: it.inspection_status === 'promoted'
                ? 'promoted'
                : 'matched',
            }
          : it
      ),
    }));
  }, []);

  const promote = useCallback(
    (itemId: string, payload: { sku: string; name: string }) => {
      setPO((cur) => ({
        ...cur,
        items: cur.items.map((it) =>
          it.id === itemId
            ? {
                ...it,
                inspection_status: 'promoted',
                promoted_sku: payload.sku,
                promoted_name: payload.name,
              }
            : it
        ),
      }));
    },
    []
  );

  const reject = useCallback((itemId: string) => {
    setPO((cur) => ({
      ...cur,
      items: cur.items.map((it) =>
        it.id === itemId ? { ...it, inspection_status: 'rejected' } : it
      ),
    }));
  }, []);

  const setStatus = useCallback((itemId: string, status: InspectionStatus) => {
    setPO((cur) => ({
      ...cur,
      items: cur.items.map((it) =>
        it.id === itemId ? { ...it, inspection_status: status } : it
      ),
    }));
  }, []);

  const totals = useMemo(() => {
    let matched = 0,
      promoted = 0,
      rejected = 0,
      pending = 0,
      totalCny = 0;
    for (const it of po.items) {
      if (it.inspection_status === 'matched') matched++;
      else if (it.inspection_status === 'promoted') promoted++;
      else if (it.inspection_status === 'rejected') rejected++;
      else pending++;
      totalCny += it.unit_price * it.quantity;
    }
    return {
      matched,
      promoted,
      rejected,
      pending,
      totalCny,
      totalKrw: Math.round(totalCny * po.exchange_rate),
    };
  }, [po]);

  return {
    po,
    selectedItemId,
    select,
    loadSample,
    reset,
    updatePO,
    updateItem,
    addItem,
    removeItem,
    addPhoto,
    removePhoto,
    completeMatch,
    promote,
    reject,
    setStatus,
    totals,
  };
}
