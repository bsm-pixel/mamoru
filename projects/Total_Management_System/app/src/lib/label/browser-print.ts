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

/** 시간 제한이 있는 조회 — Browser Print 가 프린터를 찾느라 오래 걸리면 기다리지 않는다 */
async function getJson(url: string, ms: number): Promise<any | null> {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), ms);
  try {
    const r = await fetch(url, { method: 'GET', signal: ctl.signal });
    if (!r.ok) return null;
    return await r.json();
  } catch {
    return null;
  } finally {
    clearTimeout(timer);
  }
}

/**
 * 어느 프린터로 보낼지 고른다 (2026-11-20)
 *
 * 프린터 주소는 공유기가 바꿀 수 있다(.44 → .45 → .41 …). Browser Print 의 "기본 프린터"는 옛 주소로 남아
 * 출력이 막히는 일이 되풀이됐다. 그래서 기본값만 믿지 않고, Browser Print 가 "지금 찾아 낸" 프린터를 먼저 쓴다.
 *  1) 기본 프린터가 주소 연결이고 지금도 찾아지면 → 그대로
 *  2) 아니면 지금 찾아진 주소 연결 제브라 프린터 → 그쪽으로 (주소가 바뀐 경우 · 기본값이 윈도우 드라이버인 경우)
 *  3) 찾아진 것이 없으면 → 기본 프린터 (전과 같은 동작)
 * 윈도우 드라이버 연결(connection: driver)을 뒤로 미루는 까닭: 윈도우 대기열에 넣기만 해도 "성공"으로 돌아와
 * 드라이버 포트 주소가 틀리면 아무것도 안 나오는데 화면에는 출력됐다고 뜬다.
 */
export function pickPrinter(defaultDevice: any | null, available: any[] | null): any | null {
  const network = (available || []).filter((d) => d && d.connection === 'network' && d.uid);
  if (network.length > 0) {
    if (defaultDevice?.connection === 'network') {
      const same = network.find((d) => d.uid === defaultDevice.uid);
      if (same) return same;
    }
    return network.find((d) => /zebra/i.test(String(d.manufacturer || ''))) || network[0];
  }
  return defaultDevice || null;
}

/** Browser Print 가용 + 보낼 프린터 탐색. 성공 시 {base, device} */
async function findPrinter(): Promise<{ base: string; device: any } | null> {
  for (const base of BP_BASES) {
    const def = await getJson(`${base}/default?type=printer`, 4000);
    if (!def) continue; // 이 주소로는 Browser Print 가 응답하지 않는다 → 다음 주소
    const found = await getJson(`${base}/available`, 8000);
    const device = pickPrinter(def, Array.isArray(found?.printer) ? found.printer : null);
    if (device) return { base, device };
  }
  return null;
}

/** Browser Print 로컬 서비스 + 프린터 가용 여부 */
export async function isBrowserPrintAvailable(): Promise<boolean> {
  return (await findPrinter()) !== null;
}

/** ZPL을 프린터로 전송. 성공 true / 실패(미설치·미허용 등) false */
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
