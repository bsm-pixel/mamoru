'use client';

import { use, useState, useEffect } from 'react';
import { useRouter } from 'next/navigation';
import {
  useSourcingItem, useInboundItemUpdate, useUploadInboundPhoto, useDeleteInboundPhoto,
} from '@/hooks/use-sourcing';
import { Camera, X, ExternalLink, Check, ArrowLeft, Loader2 } from 'lucide-react';
import { CameraCapture } from '@/components/ui/camera-capture';

const STATUS_LABEL: Record<string, string> = {
  pending: '대기', matched: '매칭완료', selected: '채택', rejected: '탈락',
};
const STATUS_TONE: Record<string, string> = {
  pending: 'bg-neutral-100 text-neutral-500',
  matched: 'bg-emerald-50 text-emerald-700',
  selected: 'bg-indigo-black text-white',
  rejected: 'bg-rose-50 text-rose-600',
};

/** 폰 QR 스캔 직진입 — 입고매칭 (실물 사진 촬영 + 메모 + 매칭완료) */
export default function SourcingInboundPage({ params }: { params: Promise<{ itemId: string }> }) {
  const { itemId } = use(params);
  const router = useRouter();
  const { data, isLoading, isError } = useSourcingItem(itemId);
  const update = useInboundItemUpdate(itemId);
  const upload = useUploadInboundPhoto(itemId);
  const removePhoto = useDeleteInboundPhoto(itemId);
  const [memo, setMemo] = useState('');
  const [uploading, setUploading] = useState(false);
  const [cameraOpen, setCameraOpen] = useState(false);

  const item = data?.item;
  useEffect(() => { if (item) setMemo(item.inbound_memo ?? ''); }, [item?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  if (isLoading) {
    return <div className="min-h-screen bg-neutral-50 flex items-center justify-center text-sm text-neutral-400">불러오는 중…</div>;
  }
  if (isError || !item) {
    return (
      <div className="min-h-screen bg-neutral-50 flex flex-col items-center justify-center gap-3 p-6 text-center">
        <p className="text-sm text-neutral-500">품목을 찾을 수 없습니다. (라벨이 오래됐거나 삭제됨)</p>
        <button onClick={() => router.push('/sourcing')} className="text-sm text-blue-600 underline">샘플 소싱으로</button>
      </div>
    );
  }

  const photos = item.inbound_photos ?? [];
  const seq = item.sticker_no.split('-').pop();

  const handleCaptureOne = async (file: File) => {
    if (photos.length >= 5) return;
    setUploading(true);
    try {
      await upload.mutateAsync(file);
    } finally {
      setUploading(false);
    }
  };

  const handleComplete = async () => {
    await update.mutateAsync({
      inbound_memo: memo,
      inspection_status: item.inspection_status === 'pending' ? 'matched' : item.inspection_status,
    });
  };

  return (
    <div className="min-h-screen bg-neutral-50 pb-28">
      {/* 헤더 */}
      <div className="sticky top-0 z-10 bg-white border-b border-neutral-200 px-4 py-3 flex items-center gap-3">
        <button onClick={() => router.push('/sourcing')} className="text-neutral-600"><ArrowLeft size={20} /></button>
        <div className="flex-1 min-w-0">
          <div className="text-[10px] text-neutral-400 font-mono">{item.sticker_no}</div>
          <div className="text-sm font-bold text-neutral-900">입고매칭 #{seq}</div>
        </div>
        <span className={`px-2 py-0.5 text-[10px] rounded-full font-medium ${STATUS_TONE[item.inspection_status]}`}>
          {STATUS_LABEL[item.inspection_status]}
        </span>
      </div>

      <div className="p-4 space-y-3">
        {/* 품목 정보 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="text-base font-bold text-neutral-900 leading-snug">{item.product_name || '(품목명 없음)'}</div>
          <div className="mt-1.5 text-xs text-neutral-500 flex items-center flex-wrap gap-x-3 gap-y-1">
            <span>단가 <strong className="text-neutral-900">¥{item.unit_price}</strong></span>
            {item.moq != null && <><span>·</span><span>MOQ {item.moq}</span></>}
          </div>
          {item.features_memo && (
            <div className="mt-3 px-3 py-2 bg-amber-50 border-l-2 border-amber-300 text-xs text-neutral-700 rounded">{item.features_memo}</div>
          )}
          {item.vendor_url && (
            <a href={item.vendor_url} target="_blank" rel="noreferrer" className="mt-3 inline-flex items-center gap-1 text-xs text-blue-600">
              1688 상품 페이지 <ExternalLink size={11} />
            </a>
          )}
        </div>

        {/* 사진 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="flex items-center justify-between mb-3">
            <div className="text-[11px] font-bold text-neutral-500 uppercase tracking-wider">현장 사진</div>
            <span className="text-xs text-neutral-500">{photos.length}/5</span>
          </div>
          <button
            type="button"
            onClick={() => setCameraOpen(true)}
            disabled={photos.length >= 5 || uploading}
            className="w-full py-5 bg-neutral-900 active:bg-neutral-700 text-white font-bold rounded-xl flex items-center justify-center gap-2 disabled:opacity-40"
          >
            {uploading ? <Loader2 size={20} className="animate-spin" /> : <Camera size={20} />}
            {uploading ? '업로드 중…' : photos.length === 0 ? '사진 촬영' : '추가 촬영'}
          </button>
          <CameraCapture open={cameraOpen} onClose={() => setCameraOpen(false)} onCapture={handleCaptureOne} />

          {photos.length > 0 && (
            <div className="mt-3 grid grid-cols-3 gap-2">
              {photos.map((p, i) => (
                <div key={i} className="relative aspect-square rounded-lg overflow-hidden border border-neutral-200 bg-cover bg-center" style={{ backgroundImage: `url(${p})` }}>
                  <button
                    type="button"
                    onClick={() => removePhoto.mutate(p)}
                    className="absolute top-1 right-1 w-5 h-5 bg-black/60 text-white rounded-full flex items-center justify-center active:bg-black"
                  >
                    <X size={11} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* 메모 */}
        <div className="rounded-xl bg-white p-4 shadow-sm">
          <div className="text-[11px] font-bold text-neutral-500 uppercase tracking-wider mb-2">현장 메모 (선택)</div>
          <textarea
            value={memo}
            onChange={(e) => setMemo(e.target.value)}
            placeholder="예) 손잡이 무광 처리, 흠집 1개"
            rows={3}
            className="w-full px-3 py-2 text-sm border border-neutral-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-neutral-900/10 resize-none"
          />
        </div>
      </div>

      {/* 고정 하단 CTA */}
      <div className="fixed bottom-0 left-0 right-0 bg-white border-t border-neutral-200 p-4">
        <button
          type="button"
          onClick={handleComplete}
          disabled={update.isPending}
          className="w-full py-3.5 bg-emerald-600 active:bg-emerald-700 text-white font-bold rounded-xl flex items-center justify-center gap-2 disabled:opacity-50"
        >
          {update.isPending ? <Loader2 size={18} className="animate-spin" /> : <Check size={18} />}
          매칭 완료 (저장)
        </button>
      </div>
    </div>
  );
}
