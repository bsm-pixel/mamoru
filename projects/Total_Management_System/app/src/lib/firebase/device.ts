/**
 * 기기 식별 — 푸시 구독을 "기기당 1개"로 유지하기 위한 값 (2026-09-15)
 *
 * 왜 필요한가: 전엔 사용자당 토큰 1개만 남겨서, 새 기기에서 TMS를 열면
 * 다른 기기의 토큰이 삭제됐다(PC↔모바일 알림이 서로를 끊음).
 * 이제 기기마다 고정 id를 들고 다니며 **그 기기의 옛 토큰만** 교체한다.
 */

const KEY = 'mamoru_device_id';

/** 이 브라우저(기기)의 고정 id. localStorage 에 1회 생성 후 재사용 */
export function getDeviceId(): string {
  if (typeof window === 'undefined') return '';
  try {
    let id = localStorage.getItem(KEY);
    if (!id) {
      id = (crypto?.randomUUID?.() as string) || `dev-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
      localStorage.setItem(KEY, id);
    }
    return id;
  } catch {
    // 시크릿창·저장소 차단 — 기기 식별 불가. 빈 값이면 서버가 레거시 경로로 저장한다
    return '';
  }
}

/** 설정 화면에 보여줄 사람이 읽는 기기 이름 — 예: "iPhone · Safari (앱)" */
export function getDeviceLabel(): string {
  if (typeof window === 'undefined') return '';
  const ua = navigator.userAgent || '';

  const os =
    /iPhone/i.test(ua) ? 'iPhone'
    : /iPad/i.test(ua) ? 'iPad'
    : /Android/i.test(ua) ? 'Android'
    : /Macintosh/i.test(ua) ? 'Mac'
    : /Windows/i.test(ua) ? 'Windows'
    : '기타';

  // 순서 주의: Edge·Samsung 은 UA 에 Chrome 도 들어있어 먼저 걸러야 한다
  const browser =
    /Edg\//i.test(ua) ? 'Edge'
    : /SamsungBrowser/i.test(ua) ? '삼성브라우저'
    : /CriOS|Chrome/i.test(ua) ? 'Chrome'
    : /FxiOS|Firefox/i.test(ua) ? 'Firefox'
    : /Safari/i.test(ua) ? 'Safari'
    : '브라우저';

  // PWA(홈화면 추가)로 실행 중인지 — iOS 는 PWA 여야 웹푸시가 온다
  let installed = false;
  try {
    installed =
      window.matchMedia?.('(display-mode: standalone)').matches === true ||
      (navigator as Navigator & { standalone?: boolean }).standalone === true;
  } catch { /* noop */ }

  return `${os} · ${browser}${installed ? ' (앱)' : ''}`;
}

/** iOS 는 홈화면에 추가(PWA)해야만 웹푸시가 온다 — 안내 노출 판단용 */
export function needsIosPwaNotice(): boolean {
  if (typeof window === 'undefined') return false;
  const ua = navigator.userAgent || '';
  const isIos = /iPhone|iPad|iPod/i.test(ua);
  if (!isIos) return false;
  try {
    const installed =
      window.matchMedia?.('(display-mode: standalone)').matches === true ||
      (navigator as Navigator & { standalone?: boolean }).standalone === true;
    return !installed;
  } catch {
    return true;
  }
}
