'use client';

import { useState } from 'react';
import toast from 'react-hot-toast';

interface Props {
  saleId: string;
  customerName: string;
  customerPhone: string;
  hasRepairItem: boolean;
  alreadySent: boolean;
  onClose: () => void;
  onSent: () => void;
}

export function ReviewRequestModal({ saleId, customerName, customerPhone, hasRepairItem, alreadySent, onClose, onSent }: Props) {
  const [reviewType, setReviewType] = useState<'consult' | 'repair' | 'purchase'>(hasRepairItem ? 'repair' : 'purchase');
  const [subtype, setSubtype] = useState('store_visit');
  const [sending, setSending] = useState(false);

  const handleSend = async () => {
    setSending(true);
    try {
      const res = await fetch('/api/reviews/request', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ sale_id: saleId, review_type: reviewType, subtype: reviewType === 'consult' ? subtype : undefined }),
      });
      if (!res.ok) {
        const data = await res.json();
        toast.error(data.error || '발송 실패');
        return;
      }
      toast.success('후기 요청 알림톡을 발송했습니다');
      onSent();
    } catch {
      toast.error('발송 중 오류가 발생했습니다');
    } finally {
      setSending(false);
    }
  };

  const typeOptions: Array<{ key: 'repair' | 'consult' | 'purchase'; label: string }> = [
    { key: 'repair', label: '복원수리' },
    { key: 'consult', label: '상담' },
    { key: 'purchase', label: '제품구매' },
  ];
  const subtypeOptions = [
    { key: 'store_visit', label: '매장방문' },
    { key: 'field_request', label: '출장' },
    { key: 'talk_consult', label: '톡상담' },
  ];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-[95vw] max-w-[400px] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="px-5 py-3 border-b border-neutral-200 flex items-center justify-between">
          <h3 className="text-sm font-bold text-neutral-800">후기 요청</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600 text-lg">×</button>
        </div>

        <div className="px-5 py-4 space-y-4">
          <div className="bg-neutral-50 rounded-lg px-3 py-2.5">
            <p className="text-sm font-medium">{customerName}</p>
            <p className="text-xs text-neutral-500">{customerPhone}</p>
          </div>

          <div>
            <label className="text-xs font-semibold text-neutral-600 mb-2 block">후기 유형</label>
            <div className="flex gap-2">
              {typeOptions.map((t) => (
                <button
                  key={t.key}
                  onClick={() => setReviewType(t.key)}
                  className={`flex-1 py-2 text-xs rounded-lg border transition ${
                    reviewType === t.key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                  }`}
                >{t.label}</button>
              ))}
            </div>
          </div>

          {reviewType === 'consult' && (
            <div>
              <label className="text-xs font-semibold text-neutral-600 mb-2 block">상담 방식</label>
              <div className="flex gap-2">
                {subtypeOptions.map((st) => (
                  <button
                    key={st.key}
                    onClick={() => setSubtype(st.key)}
                    className={`flex-1 py-2 text-xs rounded-lg border transition ${
                      subtype === st.key ? 'bg-neutral-900 text-white border-neutral-900' : 'bg-white text-neutral-500 border-neutral-200'
                    }`}
                  >{st.label}</button>
                ))}
              </div>
            </div>
          )}

          {alreadySent && (
            <p className="text-[11px] text-amber-600">이미 후기 요청을 보낸 건입니다. 재발송하시겠습니까?</p>
          )}
        </div>

        <div className="px-5 py-3 border-t border-neutral-200 flex gap-2">
          <button onClick={onClose} className="flex-1 py-2 rounded-lg border border-neutral-200 text-sm text-neutral-600">취소</button>
          <button
            onClick={handleSend}
            disabled={sending}
            className="flex-1 py-2 rounded-lg bg-neutral-900 text-white text-sm font-medium disabled:opacity-50"
          >
            {sending ? '발송 중...' : '알림톡 보내기'}
          </button>
        </div>
      </div>
    </div>
  );
}
