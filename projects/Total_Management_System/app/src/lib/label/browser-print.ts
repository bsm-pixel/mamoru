'use client';

/**
 * Zebra Browser Print 로컬 HTTP API로 ZPL 직접 전송 (SDK 스크립트 불필요).
 * Browser Print 앱이 PC에 설치·실행 중이면 127.0.0.1:9100 로컬 서버로 전송 → ZT231 즉시 출력.
 * 미설치/실패 시 false 반환 → 호출측에서 ZPL 다운로드 fallback.
 *
 * 참고: HTTPS(TMS) → http://127.0.0.1 은 localhost라 mixed-content 예외로 대개 허용됨.
 */

const BP_BASE = 'http://127.0.0.1:9100';

/* eslint-disable @typescript-eslint/no-explicit-any */

/** Browser Print 로컬 서비스 가용 여부 */
export async function isBrowserPrintAvailable(): Promise<boolean> {
  try {
    const r = await fetch(`${BP_BASE}/available`, { method: 'GET' });
    return r.ok;
  } catch {
    return false;
  }
}

/** 기본 프린터 device 객체 */
async function getDefaultDevice(): Promise<any | null> {
  try {
    const r = await fetch(`${BP_BASE}/default?type=printer`);
    if (!r.ok) return null;
    return await r.json();
  } catch {
    return null;
  }
}

/** ZPL을 기본 프린터로 전송. 성공 true / 실패(미설치 등) false */
export async function sendZplToPrinter(zpl: string): Promise<boolean> {
  const device = await getDefaultDevice();
  if (!device) return false;
  try {
    const r = await fetch(`${BP_BASE}/write`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ device, data: zpl }),
    });
    return r.ok;
  } catch {
    return false;
  }
}

/** ZPL 파일 다운로드 (Browser Print 미연결 시 ZSU 등으로 테스트용) */
export function downloadZpl(zpl: string, filename = 'label.zpl'): void {
  const blob = new Blob([zpl], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
