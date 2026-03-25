'use client';

import { useQuery } from '@tanstack/react-query';
import { createClient } from '@/lib/supabase/client';
import { formatPhone } from '@/lib/utils/format';
import { X, MapPin, Phone } from 'lucide-react';

interface TodayConsultationPickerProps {
  open: boolean;
  onClose: () => void;
  onSelect: (consultation: {
    id: string;
    customer_id: string | null;
    name: string;
    phone: string;
    address_road: string | null;
    address_detail: string | null;
  }) => void;
}

/** 오늘 예약 고객 목록 모달 */
export function TodayConsultationPicker({ open, onClose, onSelect }: TodayConsultationPickerProps) {
  const supabase = createClient();
  const today = new Date().toISOString().slice(0, 10);

  const { data: consultations = [], isLoading } = useQuery({
    queryKey: ['today-consultations', today],
    staleTime: 60_000,
    enabled: open,
    queryFn: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const { data, error } = await (supabase as any)
        .from('consultations')
        .select('id, customer_id, name, phone, address_road, address_detail, visit_time, consultation_type, status')
        .eq('visit_date', today)
        .in('status', ['confirmed', 'visited', 'pending'])
        .order('visit_time', { ascending: true });
      if (error) throw error;
      return (data || []) as Array<{
        id: string;
        customer_id: string | null;
        name: string;
        phone: string;
        address_road: string | null;
        address_detail: string | null;
        visit_time: string | null;
        consultation_type: string;
        status: string;
      }>;
    },
  });

  if (!open) return null;

  const TYPE_LABEL: Record<string, string> = {
    store_visit: '매장방문',
    field_request: '출장',
    talk_consult: '톡상담',
  };

  const STATUS_LABEL: Record<string, string> = {
    confirmed: '확정',
    visited: '방문',
    pending: '대기',
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40" onClick={onClose}>
      <div
        className="bg-white rounded-xl shadow-xl w-[90vw] max-w-[400px] max-h-[70vh] flex flex-col"
        onClick={(e) => e.stopPropagation()}
      >
        {/* 헤더 */}
        <div className="flex items-center justify-between px-4 py-3 border-b border-neutral-200">
          <h3 className="text-sm font-bold text-neutral-800">오늘 예약 고객</h3>
          <button onClick={onClose} className="text-neutral-400 hover:text-neutral-600">
            <X size={18} />
          </button>
        </div>

        {/* 목록 */}
        <div className="flex-1 overflow-y-auto">
          {isLoading ? (
            <div className="p-6 text-center text-sm text-neutral-400">로딩중...</div>
          ) : consultations.length === 0 ? (
            <div className="p-6 text-center text-sm text-neutral-400">오늘 예약 고객이 없습니다</div>
          ) : (
            <div className="divide-y divide-neutral-100">
              {consultations.map((c) => (
                <button
                  key={c.id}
                  onClick={() => onSelect(c)}
                  className="w-full text-left px-4 py-3 hover:bg-neutral-50 transition"
                >
                  <div className="flex items-center gap-2">
                    <span className="text-sm font-semibold text-neutral-800">{c.name}</span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-neutral-100 text-neutral-500">
                      {TYPE_LABEL[c.consultation_type] || c.consultation_type}
                    </span>
                    <span className="text-[10px] px-1.5 py-0.5 rounded bg-blue-50 text-blue-600">
                      {STATUS_LABEL[c.status] || c.status}
                    </span>
                    {c.visit_time && (
                      <span className="text-[10px] text-neutral-400">{c.visit_time}</span>
                    )}
                  </div>
                  <div className="flex items-center gap-3 mt-1 text-xs text-neutral-500">
                    <span className="flex items-center gap-1">
                      <Phone size={10} />
                      {formatPhone(c.phone)}
                    </span>
                    {c.address_road && (
                      <span className="flex items-center gap-1 truncate">
                        <MapPin size={10} />
                        {c.address_road}
                      </span>
                    )}
                  </div>
                </button>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
