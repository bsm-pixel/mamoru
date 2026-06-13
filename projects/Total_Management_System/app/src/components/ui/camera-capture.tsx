'use client';

/**
 * 인페이지 카메라 촬영 — getUserMedia 로 페이지 안에서 촬영(앱 전환 X).
 * `<input capture>` 가 모바일에서 카메라 앱을 띄우면 페이지가 evict→재로드되어 모달이 닫히고
 * 대시보드로 튕기는 문제를 원천 차단. 권한 거부/미지원이면 파일 선택 fallback.
 *
 * 사용: <CameraCapture open onCapture={(file) => ...} onClose={() => ...} />
 */

import { useEffect, useRef, useState } from 'react';
import { Camera, X, RotateCcw, ImageUp, Loader2 } from 'lucide-react';

interface Props {
  open: boolean;
  onCapture: (file: File) => void;
  onClose: () => void;
}

export function CameraCapture({ open, onCapture, onClose }: Props) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const fileRef = useRef<HTMLInputElement>(null);
  const [status, setStatus] = useState<'init' | 'live' | 'fallback' | 'error'>('init');
  const [errMsg, setErrMsg] = useState('');

  function stopStream() {
    streamRef.current?.getTracks().forEach((t) => t.stop());
    streamRef.current = null;
  }

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setStatus('init');
    setErrMsg('');

    (async () => {
      if (typeof navigator === 'undefined' || !navigator.mediaDevices?.getUserMedia) {
        if (!cancelled) setStatus('fallback');
        return;
      }
      try {
        const stream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: 'environment' } },
          audio: false,
        });
        if (cancelled) { stream.getTracks().forEach((t) => t.stop()); return; }
        streamRef.current = stream;
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          await videoRef.current.play().catch(() => {});
        }
        setStatus('live');
      } catch (e) {
        // 권한 거부 / 카메라 없음 → 파일 선택 fallback
        if (!cancelled) {
          setErrMsg(e instanceof Error ? e.message : String(e));
          setStatus('fallback');
        }
      }
    })();

    return () => { cancelled = true; stopStream(); };
  }, [open]);

  if (!open) return null;

  function handleShoot() {
    const video = videoRef.current;
    if (!video || !video.videoWidth) return;
    const canvas = document.createElement('canvas');
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (!blob) return;
      const file = new File([blob], `capture_${canvas.width}x${canvas.height}.jpg`, { type: 'image/jpeg' });
      stopStream();
      onCapture(file);
      onClose();
    }, 'image/jpeg', 0.9);
  }

  function handleClose() { stopStream(); onClose(); }

  return (
    <div className="fixed inset-0 z-[60] bg-black flex flex-col" onClick={(e) => e.stopPropagation()}>
      {/* 상단 바 */}
      <div className="flex items-center justify-between px-4 py-3 text-white shrink-0">
        <span className="text-sm font-medium">사진 촬영</span>
        <button onClick={handleClose} className="w-9 h-9 flex items-center justify-center rounded-full bg-white/10 hover:bg-white/20">
          <X size={20} />
        </button>
      </div>

      {/* 프리뷰 / 상태 */}
      <div className="flex-1 relative flex items-center justify-center overflow-hidden">
        {status === 'init' && (
          <div className="text-white/70 flex flex-col items-center gap-2"><Loader2 size={28} className="animate-spin" /><span className="text-sm">카메라 여는 중…</span></div>
        )}
        <video ref={videoRef} playsInline muted
          className={`max-h-full max-w-full ${status === 'live' ? 'block' : 'hidden'}`} style={{ objectFit: 'contain' }} />
        {status === 'fallback' && (
          <div className="text-center text-white/80 px-6">
            <Camera size={36} className="mx-auto mb-3 opacity-60" />
            <p className="text-sm mb-1">카메라를 바로 열 수 없습니다</p>
            <p className="text-xs text-white/50 mb-4">{errMsg ? '권한이 거부되었거나 카메라가 없습니다.' : '이 기기는 인페이지 촬영을 지원하지 않습니다.'}</p>
            <button onClick={() => fileRef.current?.click()}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-white text-black text-sm font-medium">
              <ImageUp size={16} /> 사진 선택 / 촬영
            </button>
          </div>
        )}
      </div>

      {/* 하단 컨트롤 */}
      <div className="shrink-0 px-4 py-5 flex items-center justify-center gap-6">
        {status === 'live' && (
          <>
            <button onClick={() => fileRef.current?.click()} title="갤러리/파일"
              className="w-11 h-11 flex items-center justify-center rounded-full bg-white/10 text-white hover:bg-white/20">
              <ImageUp size={20} />
            </button>
            <button onClick={handleShoot} aria-label="촬영"
              className="w-16 h-16 rounded-full bg-white ring-4 ring-white/30 active:scale-95 transition" />
            <button onClick={handleClose} title="취소"
              className="w-11 h-11 flex items-center justify-center rounded-full bg-white/10 text-white hover:bg-white/20">
              <RotateCcw size={20} />
            </button>
          </>
        )}
      </div>

      {/* fallback / 갤러리 파일 입력 — capture 미사용(앱 강제전환 회피) */}
      <input
        ref={fileRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          e.target.value = '';
          if (f) { stopStream(); onCapture(f); onClose(); }
        }}
      />
    </div>
  );
}
