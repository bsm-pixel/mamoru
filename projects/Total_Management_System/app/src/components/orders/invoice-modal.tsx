'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { useBookInvoice } from '@/hooks/use-orders';
import type { Order } from '@/lib/supabase/types';

interface InvoiceModalProps {
  open: boolean;
  onClose: () => void;
  order: Order;
}

export function InvoiceModal({ open, onClose, order }: InvoiceModalProps) {
  const bookInvoice = useBookInvoice();
  const [gdsNm, setGdsNm] = useState('마모루 제품');

  const handleSubmit = () => {
    bookInvoice.mutate(
      {
        orderId: order.id,
        ordNo: order.imweb_order_no,
        rcvName: order.recipient_name,
        rcvTel: order.recipient_phone || '',
        rcvZip: order.recipient_postcode || '',
        rcvAdr: `${order.recipient_address || ''} ${order.recipient_address_detail || ''}`.trim(),
        gdsNm,
        dlvMsg: order.recipient_memo || '',
      },
      { onSuccess: () => onClose() }
    );
  };

  return (
    <Modal open={open} onClose={onClose} title="송장 생성">
      <div className="space-y-4">
        {/* 수신자 정보 */}
        <div className="rounded-lg bg-warm-ivory p-4 space-y-2 text-sm">
          <div className="flex justify-between">
            <span className="text-neutral-500">수신자</span>
            <span className="font-medium">{order.recipient_name}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">전화</span>
            <span>{order.recipient_phone}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">우편번호</span>
            <span>{order.recipient_postcode}</span>
          </div>
          <div>
            <span className="text-neutral-500">주소</span>
            <p className="mt-1">{order.recipient_address} {order.recipient_address_detail}</p>
          </div>
          {order.recipient_memo && (
            <div>
              <span className="text-neutral-500">배송메모</span>
              <p className="mt-1">{order.recipient_memo}</p>
            </div>
          )}
        </div>

        {/* 상품명 */}
        <div>
          <label className="block text-sm font-medium text-neutral-700 mb-1">
            상품명
          </label>
          <input
            value={gdsNm}
            onChange={(e) => setGdsNm(e.target.value)}
            className="w-full h-10 px-3 rounded-lg border border-neutral-200 bg-warm-ivory text-indigo-black focus:outline-none focus:ring-2 focus:ring-terracotta/40 transition"
          />
        </div>

        {/* 버튼 */}
        <div className="flex gap-2 pt-2">
          <Button variant="secondary" onClick={onClose} className="flex-1">
            취소
          </Button>
          <Button
            onClick={handleSubmit}
            disabled={bookInvoice.isPending}
            className="flex-1"
          >
            {bookInvoice.isPending ? '생성 중...' : '송장 생성'}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
