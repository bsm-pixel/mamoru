'use client';

import { useSyncExternalStore } from 'react';
import { WifiOff } from 'lucide-react';

/**
 * 인터넷이 끊겼을 때 알림 (2026-11-15)
 * 전엔 끊겨도 아무 표시가 없어 저장이 안 되는 이유를 알 수 없었다. 다시 이어지면 스스로 사라진다.
 */
function subscribe(cb: () => void) {
  window.addEventListener('online', cb);
  window.addEventListener('offline', cb);
  return () => { window.removeEventListener('online', cb); window.removeEventListener('offline', cb); };
}

export function OfflineBanner() {
  const online = useSyncExternalStore(subscribe, () => navigator.onLine, () => true);
  if (online) return null;
  return (
    <div role="alert" className="fixed top-0 inset-x-0 z-[300] bg-stone-900 text-white text-xs font-semibold px-4 py-2 flex items-center justify-center gap-2">
      <WifiOff size={14} />연결이 끊겼습니다 — 지금 저장하면 실패할 수 있습니다. 인터넷이 이어지면 이 줄이 사라집니다
    </div>
  );
}
