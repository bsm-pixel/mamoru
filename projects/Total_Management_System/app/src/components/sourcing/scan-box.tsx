'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { ScanLine } from 'lucide-react';
import toast from 'react-hot-toast';

/**
 * PC 바코드/QR 스캐너 입력창.
 * 스캐너 = 키보드(HID) → 라벨 QR(입고매칭 URL) 스캔 시 URL 타이핑 + Enter.
 * URL 안의 itemId(UUID)를 추출해 해당 품목 입고매칭 페이지로 이동.
 * (매칭 후에도 동일 — 사진 있으면 보이고, 없으면 추가하는 흐름)
 */
const UUID_RE = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/;

export function ScanBox() {
  const router = useRouter();
  const [val, setVal] = useState('');

  const process = (raw: string) => {
    const text = raw.trim();
    if (!text) return;
    const m = text.match(UUID_RE);
    if (!m) {
      toast.error('소싱 라벨 QR이 아닙니다');
      setVal('');
      return;
    }
    setVal('');
    router.push(`/sourcing/inbound/${m[0]}`);
  };

  return (
    <div className="flex items-center gap-2 rounded-xl border-2 border-dashed border-indigo-black/25 bg-white px-3 py-2.5">
      <ScanLine size={18} className="text-indigo-black flex-shrink-0" />
      <input
        autoFocus
        value={val}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); process(val); } }}
        placeholder="여기 커서를 둔 채로 바코드 스캐너로 라벨 QR 스캔 → 해당 품목 자동 열림"
        className="flex-1 min-w-0 text-sm bg-transparent text-indigo-black placeholder:text-neutral-400 focus:outline-none"
      />
    </div>
  );
}
