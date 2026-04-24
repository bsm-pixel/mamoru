'use client';

/**
 * 관리자 직접 상담 등록 모달
 * 인스타 DM / 유선 접수된 건을 TMS에 수기 입력 → 기존 흐름(알림톡·리마인더·캘린더)에 자동 편입
 *
 * 지원: 매장방문 / 출장요청 (톡상담 제외)
 * 중복 감지: 같은 번호 + 같은 일시 → 경고 서브 모달 + 기존 상담 정보 표시
 */

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { Store, Truck, AlertTriangle, Info } from 'lucide-react';
import toast from 'react-hot-toast';
import {
  useCreateConsultation,
  type DuplicateExistingConsultation,
} from '@/hooks/use-consultations';

interface Props {
  open: boolean;
  onClose: () => void;
  /** 등록 성공 시 해당 상담 상세로 이동 or 선택 표시하고 싶을 때 */
  onCreated?: (consultationId: string) => void;
}

type ConsultType = 'store_visit' | 'field_request';

const TYPE_OPTIONS: { key: ConsultType; label: string; icon: React.ReactNode }[] = [
  { key: 'store_visit', label: '매장방문', icon: <Store size={14} /> },
  { key: 'field_request', label: '출장요청', icon: <Truck size={14} /> },
];

export function CreateConsultationModal({ open, onClose, onCreated }: Props) {
  const [type, setType] = useState<ConsultType>('store_visit');
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [visitDate, setVisitDate] = useState('');
  const [visitTime, setVisitTime] = useState('');
  const [addressRoad, setAddressRoad] = useState('');
  const [addressDetail, setAddressDetail] = useState('');
  const [memo, setMemo] = useState('');
  const [notify, setNotify] = useState(true);
  const [duplicate, setDuplicate] = useState<DuplicateExistingConsultation | null>(null);

  const create = useCreateConsultation();

  const reset = () => {
    setType('store_visit');
    setName('');
    setPhone('');
    setVisitDate('');
    setVisitTime('');
    setAddressRoad('');
    setAddressDetail('');
    setMemo('');
    setNotify(true);
    setDuplicate(null);
  };

  const handleClose = () => {
    if (create.isPending) return;
    reset();
    onClose();
  };

  const handleSubmit = async () => {
    // 검증
    if (!name.trim()) return toast.error('고객명을 입력해주세요');
    if (!phone.trim()) return toast.error('연락처를 입력해주세요');
    if (!/^0\d{8,10}$/.test(phone.replace(/\D/g, ''))) return toast.error('연락처 형식을 확인해주세요');
    if (!visitDate) return toast.error('날짜를 선택해주세요');
    if (!visitTime) return toast.error('시간을 선택해주세요');
    if (type === 'field_request' && !addressRoad.trim()) return toast.error('출장은 주소가 필수입니다');

    try {
      const result = await create.mutateAsync({
        type,
        name: name.trim(),
        phone: phone.trim(),
        visitDate,
        visitTime,
        addressRoad: type === 'field_request' ? addressRoad.trim() : undefined,
        addressDetail: type === 'field_request' ? addressDetail.trim() : undefined,
        memo: memo.trim() || undefined,
        notify,
      });

      if (!result.ok) {
        // 중복
        setDuplicate(result.duplicate);
        return;
      }

      const createdId = result.data.id;
      reset();
      onClose();
      if (onCreated) onCreated(createdId);
    } catch {
      // 에러 토스트는 훅에서 처리
    }
  };

  const saving = create.isPending;

  return (
    <>
      <Modal open={open && !duplicate} onClose={handleClose} title="일정 수동 등록">
        <div className="space-y-4">
          {/* 타입 선택 */}
          <div>
            <label className="block text-xs font-semibold text-neutral-700 mb-1.5">상담 유형</label>
            <div className="flex gap-2">
              {TYPE_OPTIONS.map((opt) => (
                <button
                  key={opt.key}
                  onClick={() => setType(opt.key)}
                  className={`flex-1 flex items-center justify-center gap-1.5 h-10 rounded-lg border text-sm font-semibold transition ${
                    type === opt.key
                      ? 'border-terracotta bg-terracotta/10 text-terracotta'
                      : 'border-neutral-200 bg-white text-neutral-600 hover:bg-neutral-50'
                  }`}
                >
                  {opt.icon}
                  {opt.label}
                </button>
              ))}
            </div>
          </div>

          {/* 고객명 + 연락처 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-xs font-semibold text-neutral-700 mb-1">고객명 *</label>
              <input
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="홍길동"
                className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm"
                maxLength={30}
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-neutral-700 mb-1">연락처 *</label>
              <input
                value={phone}
                onChange={(e) => setPhone(e.target.value)}
                placeholder="01012345678"
                inputMode="numeric"
                className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm font-mono"
                maxLength={13}
              />
            </div>
          </div>

          {/* 주소 (출장만) */}
          {type === 'field_request' && (
            <div className="space-y-2 rounded-lg border border-neutral-200 bg-warm-ivory p-3">
              <div className="flex items-center gap-1.5">
                <Info size={12} className="text-neutral-500" />
                <span className="text-[11px] text-neutral-500">
                  출장 등록 — 주소는 지도 표시·리마인더 문자에 사용됩니다
                </span>
              </div>
              <div>
                <label className="block text-xs font-semibold text-neutral-700 mb-1">주소 *</label>
                <input
                  value={addressRoad}
                  onChange={(e) => setAddressRoad(e.target.value)}
                  placeholder="서울 강남구 역삼동 ..."
                  className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-neutral-700 mb-1">상세 주소</label>
                <input
                  value={addressDetail}
                  onChange={(e) => setAddressDetail(e.target.value)}
                  placeholder="OO빌딩 3층 / 상호명 등"
                  className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm"
                />
              </div>
            </div>
          )}

          {/* 날짜 + 시간 */}
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="block text-xs font-semibold text-neutral-700 mb-1">방문 날짜 *</label>
              <input
                type="date"
                value={visitDate}
                onChange={(e) => setVisitDate(e.target.value)}
                className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm"
              />
            </div>
            <div>
              <label className="block text-xs font-semibold text-neutral-700 mb-1">방문 시간 *</label>
              <input
                type="time"
                value={visitTime}
                onChange={(e) => setVisitTime(e.target.value)}
                step={600}
                className="w-full h-10 px-3 rounded-lg border border-neutral-200 text-sm"
              />
            </div>
          </div>

          {/* 메모 */}
          <div>
            <label className="block text-xs font-semibold text-neutral-700 mb-1">메모 (선택)</label>
            <textarea
              value={memo}
              onChange={(e) => setMemo(e.target.value)}
              placeholder="인스타 DM으로 접수 / 유선 · 참고 사항 등"
              rows={2}
              className="w-full px-3 py-2 rounded-lg border border-neutral-200 text-sm resize-none"
              maxLength={300}
            />
          </div>

          {/* 알림톡 발송 */}
          <label className="flex items-start gap-2 p-3 rounded-lg bg-warm-ivory border border-neutral-200 cursor-pointer">
            <input
              type="checkbox"
              checked={notify}
              onChange={(e) => setNotify(e.target.checked)}
              className="mt-0.5"
            />
            <div className="flex-1">
              <div className="text-sm font-semibold text-neutral-800">고객에게 확정 알림톡 발송</div>
              <div className="text-[11px] text-neutral-500 leading-relaxed">
                {type === 'field_request'
                  ? '출장 확정 알림톡 (일정·주소·변경링크 포함)'
                  : '매장방문 확정 알림톡 (일정·변경링크 포함)'}
                <br />
                이미 구두로 안내 완료했으면 끄세요.
              </div>
            </div>
          </label>

          {/* 액션 */}
          <div className="flex justify-end gap-2 pt-2 border-t border-neutral-100">
            <Button variant="ghost" size="sm" onClick={handleClose} disabled={saving}>
              취소
            </Button>
            <Button size="sm" onClick={handleSubmit} disabled={saving} loading={saving}>
              등록
            </Button>
          </div>
        </div>
      </Modal>

      {/* 중복 경고 서브 모달 */}
      {duplicate && (
        <DuplicateWarningModal
          existing={duplicate}
          onClose={() => setDuplicate(null)}
          onViewExisting={() => {
            setDuplicate(null);
            reset();
            onClose();
            if (onCreated) onCreated(duplicate.id);
          }}
        />
      )}
    </>
  );
}

/** 중복 예약 경고 모달 */
function DuplicateWarningModal({
  existing,
  onClose,
  onViewExisting,
}: {
  existing: DuplicateExistingConsultation;
  onClose: () => void;
  onViewExisting: () => void;
}) {
  const typeLabel = existing.consultation_type === 'field_request' ? '출장요청' : '매장방문';
  const statusLabel: Record<string, string> = {
    pending_admin: '신규',
    assigned: '배정',
    suggested: '시간 제안',
    confirmed: '확정',
    reschedule_requested: '변경 요청',
    change_requested: '변경 요청',
  };

  return (
    <Modal open onClose={onClose} title="이미 약속된 일정이 있습니다">
      <div className="space-y-4">
        <div className="flex items-start gap-3 p-4 rounded-lg bg-amber-50 border border-amber-200">
          <AlertTriangle size={20} className="text-amber-600 shrink-0 mt-0.5" />
          <div className="flex-1 text-sm text-neutral-700">
            <p className="font-semibold text-amber-900 mb-1">같은 번호로 동일 일시에 예약된 건이 있습니다</p>
            <p className="text-xs text-amber-700 leading-relaxed">
              아래 기존 상담이 이미 존재합니다. 정말 동일 일정에 추가로 등록하시겠습니까?
              대부분은 중복 등록이므로 기존 건 먼저 확인해주세요.
            </p>
          </div>
        </div>

        <div className="rounded-lg border border-neutral-200 bg-white p-3 space-y-1.5 text-sm">
          <div className="flex justify-between">
            <span className="text-neutral-500">고객명</span>
            <span className="font-semibold text-neutral-900">{existing.name}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">연락처</span>
            <span className="font-mono text-neutral-700">{existing.phone}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">유형</span>
            <span className="text-neutral-700">{typeLabel}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">일시</span>
            <span className="font-semibold text-neutral-900">
              {existing.visit_date} {existing.visit_time}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-neutral-500">상태</span>
            <span className="text-neutral-700">{statusLabel[existing.status] || existing.status}</span>
          </div>
          <div className="flex justify-between pt-1 border-t border-neutral-100">
            <span className="text-neutral-400 text-xs">상담번호</span>
            <span className="text-neutral-500 text-[11px] font-mono">{existing.unique_id}</span>
          </div>
        </div>

        <div className="flex justify-end gap-2 pt-2 border-t border-neutral-100">
          <Button variant="ghost" size="sm" onClick={onClose}>
            입력 폼으로 돌아가기
          </Button>
          <Button variant="secondary" size="sm" onClick={onViewExisting}>
            기존 상담 확인
          </Button>
        </div>
      </div>
    </Modal>
  );
}
