'use client';

/**
 * 070: 판매를 출장/매장상담과 수동 연결하는 모달
 *
 * 사용 시나리오:
 *  - 기존 sale 데이터에 source_consultation_id가 없는 경우
 *  - 사장님이 "이 판매는 [상담]에서 시작된 거예요" 라고 기억할 때 수동 link
 *  - link 후 ReviewManagementCard가 자동으로 mirror 모드 전환 (중복 후기 발송 위험 0)
 *
 * 안전장치:
 *  - 같은 phone 기준만 표시 (다른 고객 잘못 묶이지 않게)
 *  - API 측에서도 phone 일치 검증 (오링크 방지)
 *  - 톡상담은 표시하지 않음 (출장/매장만 link 대상)
 */

import { useState, useEffect } from 'react';
import { X, Link2, AlertCircle } from 'lucide-react';
import toast from 'react-hot-toast';

interface ConsultationItem {
  id: string;
  unique_id: string;
  name: string;
  phone: string | null;
  consultation_type: string;
  status: string;
  visit_date: string | null;
  visit_time: string | null;
  received_at: string;
}

interface Props {
  saleId: string;
  customerPhone: string | null;
  onClose: () => void;
  onLinked: () => void;
}

const TYPE_LABEL: Record<string, string> = {
  store_visit: '매장방문',
  field_request: '출장요청',
  talk_consult: '온라인상담',
};

const STATUS_LABEL: Record<string, string> = {
  completed: '상담완료',
  in_progress: '진행중',
  confirmed: '확정',
  pending_admin: '신규접수',
  cancelled: '취소',
};

function formatMD(iso: string | null): string {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    return `${d.getMonth() + 1}-${d.getDate()}`;
  } catch { return ''; }
}

export function LinkConsultationModal({ saleId, customerPhone, onClose, onLinked }: Props) {
  const [items, setItems] = useState<ConsultationItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [linking, setLinking] = useState<string | null>(null);

  useEffect(() => {
    if (!customerPhone) {
      setItems([]);
      setLoading(false);
      return;
    }
    const phoneDigits = customerPhone.replace(/\D/g, '');
    fetch(`/api/consultation?phone=${phoneDigits}&limit=10`)
      .then((r) => r.ok ? r.json() : { consultations: [] })
      .then((d) => {
        // 톡상담은 link 대상 아님 (현장 상담/판매 흐름 외) + 취소된 상담도 제외
        const filtered = (d.consultations || []).filter((c: ConsultationItem) =>
          c.consultation_type !== 'talk_consult' && c.status !== 'cancelled'
        );
        setItems(filtered);
      })
      .catch(() => setItems([]))
      .finally(() => setLoading(false));
  }, [customerPhone]);

  const handleLink = async (consultationId: string, consultationName: string) => {
    if (linking) return;
    setLinking(consultationId);
    try {
      const res = await fetch(`/api/sales/${saleId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'link_consultation', source_consultation_id: consultationId }),
      });
      if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        toast.error(err.error || '연결 실패');
        return;
      }
      toast.success(`${consultationName}님 상담과 연결되었습니다`);
      onLinked();
    } catch (e) {
      toast.error(`오류: ${String(e)}`);
    } finally {
      setLinking(null);
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-black/50 flex items-start justify-center p-4 pt-16 overflow-y-auto" onClick={onClose}>
      <div className="bg-white rounded-xl w-full max-w-md p-5 space-y-4" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Link2 size={16} className="text-blue-600" />
            <h3 className="font-semibold text-sm">상담과 연결</h3>
          </div>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600">
            <X size={18} />
          </button>
        </div>

        <div className="text-xs text-neutral-500 leading-relaxed">
          이 판매가 시작된 상담을 선택하세요. 연결 후 후기 약속/발송은 원본 상담 한 곳에서만 관리됩니다.
        </div>

        {loading ? (
          <div className="text-center py-8 text-sm text-neutral-400">불러오는 중...</div>
        ) : !customerPhone ? (
          <div className="flex items-center gap-2 px-3 py-3 rounded-lg bg-amber-50 border border-amber-200">
            <AlertCircle size={14} className="text-amber-600 shrink-0" />
            <p className="text-xs text-amber-700">고객 전화번호가 없어 상담을 검색할 수 없습니다</p>
          </div>
        ) : items.length === 0 ? (
          <div className="text-center py-6 text-sm text-neutral-400">
            같은 전화번호의 출장/매장상담이 없습니다
          </div>
        ) : (
          <div className="space-y-1.5 max-h-[60vh] overflow-y-auto">
            {items.map((c) => (
              <button
                key={c.id}
                onClick={() => handleLink(c.id, c.name)}
                disabled={linking !== null}
                className="w-full text-left px-3 py-2.5 rounded-lg border border-neutral-200 hover:bg-blue-50 hover:border-blue-300 transition disabled:opacity-50"
              >
                <div className="flex items-center gap-2 mb-1">
                  <span className="font-mono text-xs font-bold text-neutral-700">{c.unique_id}</span>
                  <span className="text-[10px] px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-600">{TYPE_LABEL[c.consultation_type] || c.consultation_type}</span>
                  <span className="text-[10px] px-1.5 py-0.5 rounded bg-neutral-50 text-neutral-500">{STATUS_LABEL[c.status] || c.status}</span>
                </div>
                <div className="flex items-center justify-between text-xs">
                  <span className="text-neutral-700">{c.name}</span>
                  <span className="text-neutral-400">
                    {c.visit_date ? `${formatMD(c.visit_date)}${c.visit_time ? ` ${c.visit_time}` : ''}` : `접수 ${formatMD(c.received_at)}`}
                  </span>
                </div>
                {linking === c.id && <p className="text-[10px] text-blue-600 mt-1">연결 중...</p>}
              </button>
            ))}
          </div>
        )}

        <div className="text-[10px] text-neutral-400 pt-2 border-t border-neutral-100">
          ⚠️ 같은 고객의 다른 출장/매장상담만 표시됩니다. 톡상담·취소된 건은 제외.
        </div>
      </div>
    </div>
  );
}
