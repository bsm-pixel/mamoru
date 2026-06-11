'use client';

/**
 * Zebra Browser Print 로컬 API로 ZPL 직접 전송 (SDK 스크립트 불필요).
 * Browser Print 앱이 PC에 설치·실행 중이면 로컬 서버로 전송 → ZT231 즉시 출력.
 *
 * HTTPS(TMS)에선 HTTPS 엔드포인트 `https://localhost:9101` 사용(인증서 CN=localhost라 127.0.0.1 아님).
 * 미설치/미허용/실패 시 false → 호출측 ZPL 다운로드 fallback.
 * ⚠️ Browser Print 설정의 "Accepted Hosts"에 TMS 도메인(app-eta-sandy-75.vercel.app)이 있어야 허용됨.
 */

/* eslint-disable @typescript-eslint/no-explicit-any */

// 우선순위: HTTPS 9101(HTTPS 페이지용) → HTTP 9100(구버전/HTTP)
const BP_BASES = ['https://localhost:9101', 'http://localhost:9100', 'http://127.0.0.1:9100'];

async function fetchDefaultDevice(base: string): Promise<any | null> {
  try {
    const r = await fetch(`${base}/default?type=printer`, { method: 'GET' });
    if (!r.ok) return null;
    return await r.json();
  } catch {
    return null;
  }
}

/** Browser Print 가용 + 기본 프린터 탐색. 성공 시 {base, device} */
async function findPrinter(): Promise<{ base: string; device: any } | null> {
  for (const base of BP_BASES) {
    const device = await fetchDefaultDevice(base);
    if (device) return { base, device };
  }
  return null;
}

/** Browser Print 로컬 서비스 + 프린터 가용 여부 */
export async function isBrowserPrintAvailable(): Promise<boolean> {
  return (await findPrinter()) !== null;
}

/** ZPL을 기본 프린터로 전송. 성공 true / 실패(미설치·미허용 등) false */
export async function sendZplToPrinter(zpl: string): Promise<boolean> {
  const found = await findPrinter();
  if (!found) return false;
  try {
    const r = await fetch(`${found.base}/write`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ device: found.device, data: zpl }),
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
