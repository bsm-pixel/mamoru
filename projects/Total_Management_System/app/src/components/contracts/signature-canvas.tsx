'use client';

import { useRef, useEffect, useState, useCallback } from 'react';
import { Button } from '@/components/ui/button';
import { RotateCcw } from 'lucide-react';

interface SignatureCanvasProps {
  onSign: (dataUrl: string) => void;
  width?: number;
  height?: number;
}

export function SignatureCanvas({ onSign, width = 400, height = 200 }: SignatureCanvasProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasDrawn, setHasDrawn] = useState(false);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    // 고해상도 대응
    const dpr = window.devicePixelRatio || 1;
    canvas.width = width * dpr;
    canvas.height = height * dpr;
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;
    ctx.scale(dpr, dpr);
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.strokeStyle = '#181725';
  }, [width, height]);

  // 2026-11-23 터치·마우스 이벤트를 섞어 쓰다 손가락을 뗀 뒤 선이 칸 끝까지 이어지던 문제 → 포인터 이벤트 하나로.
  //   손가락을 떼거나 끊기거나 칸 밖으로 나가면 그 획을 끝낸다. (계약서 작성 화면은 sign-pad.tsx 의 전체 화면 서명판을 쓴다)
  const getPos = useCallback((e: React.PointerEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }, []);

  const startDraw = useCallback((e: React.PointerEvent) => {
    e.preventDefault();
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    setIsDrawing(true);
    const pos = getPos(e);
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
  }, [getPos]);

  const draw = useCallback((e: React.PointerEvent) => {
    if (!isDrawing) return;
    e.preventDefault();
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    const pos = getPos(e);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    setHasDrawn(true);
  }, [isDrawing, getPos]);

  const endDraw = useCallback(() => {
    if (!isDrawing) return;
    setIsDrawing(false);
    canvasRef.current?.getContext('2d')?.beginPath(); // 다음 획이 이전 획 끝에서 이어지지 않게
    if (hasDrawn && canvasRef.current) {
      onSign(canvasRef.current.toDataURL('image/png'));
    }
  }, [isDrawing, hasDrawn, onSign]);

  const clear = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    ctx.clearRect(0, 0, canvas.width / dpr, canvas.height / dpr);
    setHasDrawn(false);
    onSign('');
  }, [onSign]);

  return (
    <div className="space-y-2">
      <div className="relative border-2 border-dashed border-neutral-300 rounded-lg bg-white overflow-hidden">
        <canvas
          ref={canvasRef}
          className="cursor-crosshair touch-none"
          onPointerDown={startDraw}
          onPointerMove={draw}
          onPointerUp={endDraw}
          onPointerCancel={endDraw}
          onPointerLeave={endDraw}
        />
        {!hasDrawn && (
          <p className="absolute inset-0 flex items-center justify-center text-sm text-neutral-400 pointer-events-none">
            여기에 서명해주세요
          </p>
        )}
      </div>
      <Button variant="ghost" size="sm" onClick={clear} disabled={!hasDrawn}>
        <RotateCcw size={14} />
        지우기
      </Button>
    </div>
  );
}
