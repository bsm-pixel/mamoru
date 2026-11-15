'use client';

import { useEffect, useState } from 'react';
import { ConfirmModal } from './confirm-modal';

/**
 * 확인 창 하나로 (2026-11-15) — 브라우저 기본 confirm() 대신 쓴다
 *
 * 전엔 소싱 · 부자재 · 경비 등 40여 곳이 브라우저 기본 창(confirm)을, 제품 · 발주 · 판매는 자체 창을 써서
 * 같은 "삭제할까요?"가 화면마다 다르게 생겼다. 이제 어디서든:
 *
 *   if (!(await confirmAsk('이 발주를 영구 삭제합니다. 되돌릴 수 없습니다.'))) return;
 *
 * - 첫 줄이 제목, 나머지가 설명. 삭제 · 취소 · 해제 · 초기화가 들어 있으면 빨간 버튼.
 * - 버튼 이름은 문구에서 고른다(삭제 / 취소 확정 / 해제 / 확인). 바꾸려면 opts.confirmLabel.
 * - 닫는 쪽 버튼은 항상 「닫기」, 커서도 그쪽에 둔다(ConfirmModal).
 * <ConfirmAskHost /> 는 app/layout.tsx 에 한 번만 둔다.
 */
export interface ConfirmAskOptions {
  title?: string;
  confirmLabel?: string;
  danger?: boolean;
}

interface Pending { message: string; opts: ConfirmAskOptions; resolve: (ok: boolean) => void }

let push: ((p: Pending) => void) | null = null;

export function confirmAsk(message: string, opts: ConfirmAskOptions = {}): Promise<boolean> {
  return new Promise((resolve) => {
    // 호스트가 아직 없으면(아주 이른 시점) 브라우저 기본 창으로라도 묻는다 — 묻지 않고 진행하는 일은 없게
    if (!push) { resolve(typeof window !== 'undefined' ? window.confirm(message) : false); return; }
    push({ message, opts, resolve });
  });
}

function pickLabel(text: string): { label: string; danger: boolean } {
  if (/삭제|지웁|지울|지우/.test(text)) return { label: '삭제', danger: true };
  if (/초기화/.test(text)) return { label: '초기화', danger: true };
  if (/해제/.test(text)) return { label: '해제', danger: true };
  if (/취소/.test(text)) return { label: '취소 확정', danger: true };
  return { label: '확인', danger: false };
}

export function ConfirmAskHost() {
  const [queue, setQueue] = useState<Pending[]>([]);
  useEffect(() => {
    push = (p) => setQueue((q) => [...q, p]);
    return () => { push = null; };
  }, []);

  const cur = queue[0];
  if (!cur) return null;
  const done = (ok: boolean) => { cur.resolve(ok); setQueue((q) => q.slice(1)); };

  const lines = cur.message.split('\n');
  const firstLine = lines[0].trim();
  const rest = lines.slice(1).join('\n').trim();
  const picked = pickLabel(cur.message);
  // 첫 줄이 짧으면 제목으로, 길면 제목은 「확인」으로 두고 전체를 본문에
  const useFirstAsTitle = !cur.opts.title && firstLine.length <= 28 && rest.length > 0;
  const title = cur.opts.title || (useFirstAsTitle ? firstLine : picked.danger ? `${picked.label} 확인` : '확인');
  const body = useFirstAsTitle ? rest : cur.message;

  return (
    <ConfirmModal
      open
      onClose={() => done(false)}
      onConfirm={() => { cur.resolve(true); setQueue((q) => q.slice(1)); }}
      title={title}
      message={<span className="whitespace-pre-line break-keep">{body}</span>}
      confirmLabel={cur.opts.confirmLabel || picked.label}
      variant={(cur.opts.danger ?? picked.danger) ? 'danger' : 'default'}
    />
  );
}
