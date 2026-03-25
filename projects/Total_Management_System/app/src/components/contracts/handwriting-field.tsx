'use client';

import { useRef, useEffect, useState, useCallback } from 'react';

interface HandwritingFieldProps {
  label: string;
  /** 자동기입 시 배경에 표시할 텍스트 */
  placeholder?: string;
  /** 캔버스 높이 (기본 50) */
  height?: number;
  onDraw: (dataUrl: string) => void;
}

/**
 * 밑줄 위 필기 캔버스.
 * S펜/터치/마우스 지원. placeholder가 있으면 연한 배경 텍스트로 표시.
 */
export function HandwritingField({ label, placeholder, height = 50, onDraw }: HandwritingFieldProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasDrawn, setHasDrawn] = useState(false);
  const [canvasWidth, setCanvasWidth] = useState(400);

  // 컨테이너 너비에 맞춰 캔버스 리사이즈
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    const ro = new ResizeObserver((entries) => {
      const w = entries[0]?.contentRect.width;
      if (w && w > 0) setCanvasWidth(Math.floor(w));
    });
    ro.observe(container);
    return () => ro.disconnect();
  }, []);

  // 캔버스 초기화 + 배경 텍스트
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const dpr = window.devicePixelRatio || 1;
    canvas.width = canvasWidth * dpr;
    canvas.height = height * dpr;
    canvas.style.width = `${canvasWidth}px`;
    canvas.style.height = `${height}px`;
    ctx.scale(dpr, dpr);

    // 배경 텍스트 (placeholder)
    if (placeholder && !hasDrawn) {
      ctx.fillStyle = '#d4d4d4';
      ctx.font = `${Math.min(height * 0.5, 18)}px "Noto Sans KR", sans-serif`;
      ctx.textBaseline = 'middle';
      ctx.fillText(placeholder, 4, height / 2);
    }

    // 밑줄
    ctx.strokeStyle = '#d4d4d4';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, height - 1);
    ctx.lineTo(canvasWidth, height - 1);
    ctx.stroke();

    // 드로잉 스타일 복원
    ctx.strokeStyle = '#181725';
    ctx.lineWidth = 1.8;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
  }, [canvasWidth, height, placeholder, hasDrawn]);

  const getPos = useCallback((e: React.MouseEvent | React.TouchEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return { x: 0, y: 0 };
    const rect = canvas.getBoundingClientRect();
    if ('touches' in e) {
      const touch = e.touches[0];
      return { x: touch.clientX - rect.left, y: touch.clientY - rect.top };
    }
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }, []);

  const startDraw = useCallback((e: React.MouseEvent | React.TouchEvent) => {
    e.preventDefault();
    setIsDrawing(true);
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    // 첫 획 시 배경 텍스트 지우고 밑줄만 유지
    if (!hasDrawn) {
      const dpr = window.devicePixelRatio || 1;
      ctx.clearRect(0, 0, canvasRef.current!.width / dpr, canvasRef.current!.height / dpr);
      // 밑줄 재그리기
      ctx.strokeStyle = '#d4d4d4';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(0, height - 1);
      ctx.lineTo(canvasWidth, height - 1);
      ctx.stroke();
      // 드로잉 스타일 복원
      ctx.strokeStyle = '#181725';
      ctx.lineWidth = 1.8;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
    }
    const pos = getPos(e);
    ctx.beginPath();
    ctx.moveTo(pos.x, pos.y);
  }, [getPos, hasDrawn, height, canvasWidth]);

  const draw = useCallback((e: React.MouseEvent | React.TouchEvent) => {
    if (!isDrawing) return;
    e.preventDefault();
    const ctx = canvasRef.current?.getContext('2d');
    if (!ctx) return;
    const pos = getPos(e);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    if (!hasDrawn) setHasDrawn(true);
  }, [isDrawing, getPos, hasDrawn]);

  const endDraw = useCallback(() => {
    setIsDrawing(false);
    if (hasDrawn && canvasRef.current) {
      onDraw(canvasRef.current.toDataURL('image/png'));
    }
  }, [hasDrawn, onDraw]);

  const clear = useCallback(() => {
    setHasDrawn(false);
    onDraw('');
  }, [onDraw]);

  return (
    <div ref={containerRef} className="w-full">
      <div className="flex items-center justify-between mb-1">
        <label className="text-[10px] text-neutral-500">{label}</label>
        {hasDrawn && (
          <button onClick={clear} className="text-[10px] text-neutral-400 hover:text-neutral-600 underline">
            지우기
          </button>
        )}
      </div>
      <canvas
        ref={canvasRef}
        className="cursor-crosshair touch-none w-full"
        onMouseDown={startDraw}
        onMouseMove={draw}
        onMouseUp={endDraw}
        onMouseLeave={endDraw}
        onTouchStart={startDraw}
        onTouchMove={draw}
        onTouchEnd={endDraw}
      />
    </div>
  );
}
