'use client';

import toast from 'react-hot-toast';

/**
 * 5초 [되돌리기] (2026-11-21)
 *
 * 두 가지 쓰임:
 *  ① undoToast  — 이미 실행한 일을 되돌릴 길이 있을 때(예: 메모 삭제 → 되살리기)
 *  ② runLater   — 실행하면 못 되돌리는 일을 5초 미뤘다가 실행(예: 반품 취소 · 발주 취소는 끝 상태라 돌아갈 길이 없다)
 *
 * 붙이면 안 되는 곳: 알림톡이 나가거나(복원수리 · 상담 취소), 택배사 · 아임웹 · 장부가 함께 바뀌는 취소(판매 · 주문 · 송장).
 * 그런 일은 미뤄도 고객 · 바깥 시스템 쪽 시각이 어긋나므로 지금처럼 확인창 한 번으로 둔다.
 */
const SECONDS = 5;

function UndoBody({ message, onUndo, id }: { message: string; onUndo: () => void; id: string }) {
  return (
    <span className="flex items-center gap-3">
      <span className="text-sm">{message}</span>
      <button
        type="button"
        data-undo
        onClick={() => { toast.dismiss(id); onUndo(); }}
        className="shrink-0 rounded-md border border-neutral-300 px-2 py-1 text-xs font-bold text-neutral-800 hover:bg-neutral-100"
      >
        되돌리기
      </button>
    </span>
  );
}

/** ① 이미 실행한 일 — 5초 동안 [되돌리기] 를 보여 준다 */
export function undoToast(message: string, onUndo: () => void): void {
  toast((t) => <UndoBody message={message} onUndo={onUndo} id={t.id} />, { duration: SECONDS * 1000 });
}

// 미뤄 둔 일 — 화면을 떠나거나 탭을 닫으면 기다리지 않고 바로 실행한다(빠뜨리지 않게)
const pending = new Map<string, () => void>();
let hooked = false;
function flushAll() { for (const [id, fire] of Array.from(pending)) { pending.delete(id); toast.dismiss(id); fire(); } }
function hookOnce() {
  if (hooked || typeof window === 'undefined') return;
  hooked = true;
  window.addEventListener('pagehide', flushAll);
}

/** ② 5초 뒤에 실행 — 그 사이 [되돌리기] 를 누르면 실행하지 않는다 */
export function runLater(message: string, run: () => void): void {
  hookOnce();
  let done = false;
  const id = toast((t) => (
    <UndoBody message={message} id={t.id} onUndo={() => { done = true; pending.delete(t.id); toast('되돌렸습니다', { duration: 1500 }); }} />
  ), { duration: SECONDS * 1000 });
  const fire = () => { if (done) return; done = true; pending.delete(id); run(); };
  pending.set(id, fire);
  setTimeout(fire, SECONDS * 1000);
}

/** 다른 화면으로 넘어가기 직전에 부른다(같은 탭 안 이동은 pagehide 가 나지 않는다) */
export function flushPendingActions(): void { flushAll(); }
