'use client';

import { useRef, useState } from 'react';
import { useQueryClient } from '@tanstack/react-query';
import toast from 'react-hot-toast';
import { ImagePlus, Star, X, ChevronLeft, ChevronRight, Loader2 } from 'lucide-react';

/**
 * 재고판매(LS) 상세 이미지 매니저 — 2026-07-21
 * 여러 장 업로드 · 대표컷 지정(첫 장) · 순서 이동 · 삭제.
 * products.tags.images 를 /api/products/[id]/images 로 관리. 대표 = image_url.
 */
export function StockImageManager({ productId, images }: { productId: string; images: string[] }) {
  const qc = useQueryClient();
  const fileRef = useRef<HTMLInputElement>(null);
  const [busy, setBusy] = useState(false);

  const refresh = () => {
    qc.invalidateQueries({ queryKey: ['product', productId] });
    qc.invalidateQueries({ queryKey: ['products'] });
  };

  const call = async (fn: () => Promise<Response>, okMsg?: string) => {
    setBusy(true);
    try {
      const res = await fn();
      const d = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(d.error || '실패');
      if (okMsg) toast.success(okMsg);
      refresh();
    } catch (e) {
      toast.error(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const onPick = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    const fd = new FormData();
    files.forEach((f) => fd.append('file', f));
    call(() => fetch(`/api/products/${productId}/images`, { method: 'POST', body: fd }), '이미지를 올렸습니다');
    if (fileRef.current) fileRef.current.value = '';
  };

  const remove = (url: string) =>
    call(() => fetch(`/api/products/${productId}/images?url=${encodeURIComponent(url)}`, { method: 'DELETE' }));

  const reorder = (from: number, to: number) => {
    if (to < 0 || to >= images.length) return;
    const next = [...images];
    const [m] = next.splice(from, 1);
    next.splice(to, 0, m);
    call(() => fetch(`/api/products/${productId}/images`, {
      method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ images: next }),
    }));
  };

  return (
    <div className="rounded-lg border border-amber-200 bg-amber-50/40 p-3">
      <div className="flex items-center justify-between mb-2">
        <h4 className="text-xs font-bold text-amber-800">상세 이미지 <span className="font-normal text-amber-600">고객 상세 모달용 · 최대 8장</span></h4>
        <button
          onClick={() => fileRef.current?.click()}
          disabled={busy || images.length >= 8}
          className="flex items-center gap-1 text-[11px] font-semibold text-white bg-neutral-900 px-2.5 py-1 rounded-lg disabled:opacity-40"
        >
          {busy ? <Loader2 size={12} className="animate-spin" /> : <ImagePlus size={12} />} 사진 추가
        </button>
        <input ref={fileRef} type="file" accept="image/*" multiple hidden onChange={onPick} />
      </div>

      {images.length === 0 ? (
        <p className="text-[11px] text-neutral-400 py-3 text-center">
          아직 이미지가 없습니다. 첫 장이 <b>대표 썸네일</b>이 됩니다.
        </p>
      ) : (
        <div className="grid grid-cols-4 gap-2">
          {images.map((url, i) => (
            <div key={url} className="relative group rounded-lg overflow-hidden border border-neutral-200 bg-white aspect-square">
              {/* eslint-disable-next-line @next/next/no-img-element */}
              <img src={url} alt={`상세 ${i + 1}`} className="w-full h-full object-cover" />
              {i === 0 && (
                <span className="absolute top-1 left-1 flex items-center gap-0.5 text-[9px] font-bold text-white bg-neutral-900/80 px-1.5 py-0.5 rounded-full">
                  <Star size={8} className="fill-white" /> 대표
                </span>
              )}
              {/* 컨트롤 */}
              <div className="absolute inset-x-0 bottom-0 flex items-center justify-center gap-1 bg-black/45 py-1 opacity-0 group-hover:opacity-100 transition">
                <button onClick={() => reorder(i, i - 1)} disabled={busy || i === 0} className="text-white disabled:opacity-30" title="앞으로"><ChevronLeft size={13} /></button>
                {i !== 0 && (
                  <button onClick={() => reorder(i, 0)} disabled={busy} className="text-white" title="대표로"><Star size={12} /></button>
                )}
                <button onClick={() => reorder(i, i + 1)} disabled={busy || i === images.length - 1} className="text-white disabled:opacity-30" title="뒤로"><ChevronRight size={13} /></button>
              </div>
              <button
                onClick={() => remove(url)}
                disabled={busy}
                className="absolute top-1 right-1 w-5 h-5 flex items-center justify-center rounded-full bg-black/55 text-white hover:bg-red-500"
                title="삭제"
              ><X size={11} /></button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
