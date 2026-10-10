'use client';

import { useState } from 'react';
import { Modal } from '@/components/ui/modal';
import { Button } from '@/components/ui/button';
import { DaumPostcodeButton } from '@/components/shared/daum-postcode-button';
import { MapPin } from 'lucide-react';

export interface ShipAddress { postcode: string; road: string; detail: string }
type AddrLike = { postcode?: string | null; road?: string | null; detail?: string | null } | null | undefined;

interface Props {
  open: boolean;
  onClose: () => void;
  /** 받는 분 — 바꾸지 않는다(사장님 결정). 보여주기만 */
  name: string;
  phone?: string | null;
  /** 고객·거래처에 등록된 주소. 없으면 별도 배송지 입력이 강제된다 */
  registered: AddrLike;
  /** 이 건에 이미 쓴 별도 배송지 — 재발급 때 그대로 미리 채운다 */
  saved?: AddrLike;
  busy?: boolean;
  confirmLabel?: string;
  /** null = 등록 주소로 발송 / 객체 = 그 주소로 발송 */
  onConfirm: (addr: ShipAddress | null) => void;
}

const full = (a: AddrLike) => [a?.road, a?.detail].filter(Boolean).join(' ');
const has = (a: AddrLike) => !!(a?.postcode && a?.road);

/**
 * 송장 생성 전 「받는 곳」 확인 모달 (159, 2026-10-10)
 *
 * 사장님 요청: 거래처·고객의 등록 주소가 아니라 그때그때 다른 곳으로 보내는 일이 잦다.
 *   전엔 고객정보 주소를 고쳤다가 되돌려야 했다(고객정보 오염 + 어디로 보냈는지 기록 없음).
 *   이제 이 모달에서 이번 건만 쓸 주소를 검색해 넣는다. 받는 사람 이름·연락처는 그대로 둔다.
 *
 * 판매(B2C)·납품(B2B) 두 화면이 같이 쓴다.
 */
export function ShipAddressModal({
  open, onClose, name, phone, registered, saved, busy, confirmLabel = '이 주소로 송장 생성', onConfirm,
}: Props) {
  const regOk = has(registered);
  // 처음 그릴 때부터 올바른 값으로 — 지난번 별도 배송지가 있으면 그걸로, 등록 주소가 없으면 입력 모드
  const initial = () => {
    const start = has(saved) ? saved : null;
    return {
      other: !!start || !regOk,
      addr: { postcode: start?.postcode || '', road: start?.road || '', detail: start?.detail || '' },
    };
  };
  const [useOther, setUseOther] = useState(() => initial().other);
  const [addr, setAddr] = useState<ShipAddress>(() => initial().addr);

  // 열릴 때 초기화 — 지난번 별도 배송지가 있으면 그걸로, 없으면 등록 주소 선택.
  //   effect 가 아니라 렌더 중 보정(React 공식 패턴) — effect 로 하면 한 번 더 그려진다
  const [wasOpen, setWasOpen] = useState(open);
  if (open !== wasOpen) {
    setWasOpen(open);
    if (open) {
      const next = initial();
      setUseOther(next.other);
      setAddr(next.addr);
    }
  }

  const ready = useOther ? !!(addr.postcode && addr.road) : regOk;

  return (
    <Modal open={open} onClose={onClose} title="택배 발송 — 받는 곳 확인">
      <div className="space-y-4">
        {/* 받는 분 — 고정 */}
        <div className="rounded-lg bg-stone-50 px-3 py-2.5 text-sm">
          <div className="flex items-center justify-between">
            <span className="text-neutral-500">받는 분</span>
            <span className="font-semibold text-neutral-900">
              {name}{phone ? <span className="ml-1.5 font-normal text-neutral-500">{phone}</span> : null}
            </span>
          </div>
        </div>

        {/* 어디로 보낼지 */}
        <div className="space-y-2">
          <button
            type="button"
            onClick={() => setUseOther(false)}
            disabled={!regOk}
            className={`w-full text-left rounded-lg border px-3 py-2.5 transition disabled:opacity-40 disabled:cursor-not-allowed ${
              !useOther && regOk ? 'border-neutral-900 bg-white' : 'border-neutral-200 bg-white hover:bg-stone-50'
            }`}
          >
            <span className="block text-sm font-semibold text-neutral-900">등록된 주소로 보내기</span>
            <span className="block text-xs text-neutral-500 mt-0.5 break-keep">
              {regOk ? `(${registered?.postcode}) ${full(registered)}` : '등록된 주소가 없습니다 — 아래에서 입력해주세요'}
            </span>
          </button>

          <button
            type="button"
            onClick={() => setUseOther(true)}
            className={`w-full text-left rounded-lg border px-3 py-2.5 transition ${
              useOther ? 'border-neutral-900 bg-white' : 'border-neutral-200 bg-white hover:bg-stone-50'
            }`}
          >
            <span className="block text-sm font-semibold text-neutral-900">이번만 다른 주소로 보내기</span>
            <span className="block text-xs text-neutral-500 mt-0.5">고객 정보의 주소는 바뀌지 않습니다</span>
          </button>
        </div>

        {/* 별도 배송지 입력 */}
        {useOther && (
          <div className="space-y-2 rounded-lg border border-neutral-200 p-3">
            <div className="flex gap-2">
              <input
                value={addr.postcode}
                readOnly
                placeholder="우편번호"
                className="w-24 h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm text-neutral-600"
              />
              <DaumPostcodeButton
                onSelected={(d) => setAddr((p) => ({ ...p, postcode: d.zonecode, road: d.roadAddress, detail: '' }))}
              >
                주소 검색
              </DaumPostcodeButton>
            </div>
            <input
              value={addr.road}
              readOnly
              placeholder="도로명 주소 (주소 검색으로 입력)"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 bg-stone-50 text-sm text-neutral-600"
            />
            <input
              value={addr.detail}
              onChange={(e) => setAddr((p) => ({ ...p, detail: e.target.value }))}
              placeholder="상세 주소 (동·호수, 매장명 등)"
              className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm"
            />
          </div>
        )}

        {/* 최종 확인 — 어디로 가는지 한 줄 */}
        <p className="flex items-start gap-1.5 text-xs text-neutral-500 break-keep">
          <MapPin size={13} className="mt-0.5 shrink-0 text-neutral-400" />
          {useOther
            ? (addr.road ? `${addr.postcode} ${[addr.road, addr.detail].filter(Boolean).join(' ')} 로 발송합니다` : '주소를 검색해주세요')
            : (regOk ? '등록된 주소로 발송합니다' : '보낼 주소가 없습니다')}
        </p>

        <div className="flex gap-2 pt-1">
          <Button variant="secondary" className="flex-1" onClick={onClose} disabled={busy}>취소</Button>
          <Button
            className="flex-1"
            disabled={!ready || busy}
            loading={busy}
            onClick={() => onConfirm(useOther ? { ...addr, detail: addr.detail.trim() } : null)}
          >
            {confirmLabel}
          </Button>
        </div>
      </div>
    </Modal>
  );
}
