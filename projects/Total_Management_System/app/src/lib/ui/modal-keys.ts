/**
 * 입력 창 공통 키보드 동작 (2026-11-15) — 창 61곳(손수 짠 배경 37 + 공용 Modal 24)이 같이 쓴다.
 *
 *  ① 열리면 커서가 첫 입력 칸으로 (입력 칸이 없으면 안전한 쪽 [닫기]/[취소] 버튼으로)
 *  ② 닫히면 연 버튼으로 돌아간다
 *  ③ Tab 은 창 안에서만 돈다 (공용 Modal 은 브라우저 <dialog> 가 이미 해 줌)
 *  ④ Ctrl+Enter = 저장(창의 주 버튼) / 한 줄 칸에서 Enter = 다음 칸
 *  ⑤ 사용자가 무언가 입력했는지 기억 → Esc 로 닫을 때 "버릴까요?" 를 물을 수 있게
 *
 * 기능·저장 흐름은 건드리지 않는다. 버튼을 대신 눌러 줄 뿐이다.
 */

const FIELD = 'input:not([type=hidden]):not([disabled]):not([readonly]), textarea:not([disabled]):not([readonly]), select:not([disabled])';
const FOCUSABLE = `${FIELD}, button:not([disabled]), a[href], [tabindex]:not([tabindex="-1"])`;
const SAFE_LABEL = /^(닫기|취소|아니오|아니요|계속 입력|돌아가기)/;
const PRIMARY_CLASS = /(^|\s)(bg-terracotta|bg-stone-900|bg-neutral-900|bg-indigo-black|bg-black)(\s|$)/;

const visible = (el: HTMLElement) => el.offsetParent !== null || el.getClientRects().length > 0;
const list = (root: HTMLElement, sel: string) =>
  Array.from(root.querySelectorAll<HTMLElement>(sel)).filter((el) => visible(el) && !el.closest('[data-modal-ignore]'));

/** 검색 칸·고르기 칸처럼 "입력한 내용"으로 치지 않는 칸 */
function isScratch(el: Element): boolean {
  const i = el as HTMLInputElement;
  return i.type === 'search' || i.type === 'file' || /검색|찾기/.test(i.placeholder || '') || !!el.closest('[data-nodirty]');
}

/** 이 칸에 자체 키 처리(검색 목록 고르기 등)가 걸려 있는가 — 있으면 Enter 를 가로채지 않는다 */
function hasOwnKeyHandler(el: Element): boolean {
  const key = Object.keys(el).find((k) => k.startsWith('__reactProps$'));
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const props = key ? (el as any)[key] : null;
  return !!(props && (props.onKeyDown || props.onKeyUp || props.onKeyPress));
}

/** 창의 주 버튼(저장·등록·확인). 삭제 같은 빨간 버튼은 주 버튼으로 보지 않는다 */
export function findPrimaryButton(root: HTMLElement): HTMLButtonElement | null {
  const marked = list(root, '[data-primary]:not([disabled])')[0] as HTMLButtonElement | undefined;
  if (marked) return marked;
  const submit = list(root, 'button[type=submit]:not([disabled])').pop() as HTMLButtonElement | undefined;
  if (submit) return submit;
  const dark = (list(root, 'button:not([disabled])') as HTMLButtonElement[]).filter((b) => PRIMARY_CLASS.test(b.className));
  return dark.length ? dark[dark.length - 1] : null;
}

export interface ModalKeys {
  /** 사용자가 창 안에서 무언가 입력·선택했는가 */
  isDirty: () => boolean;
  /** 뒷정리 + 연 버튼으로 초점 되돌리기 */
  dispose: () => void;
}

export function attachModalKeys(root: HTMLElement, opts: { trapTab: boolean; isTop: () => boolean }): ModalKeys {
  const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  let dirty = false;

  // ① 첫 초점 — 이미 창 안 입력 칸에 초점이 있으면(autoFocus) 그대로 둔다
  const placeFocus = () => {
    const cur = document.activeElement;
    if (cur instanceof HTMLElement && root.contains(cur) && cur.matches(FIELD)) return;
    // 폰·태블릿(손가락)에서는 입력 칸에 자동으로 초점을 주지 않는다 — 창이 열리자마자 자판이 화면을 가린다
    const touch = typeof window.matchMedia === 'function' && window.matchMedia('(pointer: coarse)').matches;
    const field = touch ? undefined : list(root, FIELD).find((el) => (el as HTMLInputElement).type !== 'checkbox' && (el as HTMLInputElement).type !== 'radio' && (el as HTMLInputElement).type !== 'file');
    if (field) { field.focus({ preventScroll: true }); return; }
    const buttons = list(root, 'button:not([disabled])');
    const safe = buttons.find((b) => b.hasAttribute('data-safe')) || buttons.find((b) => SAFE_LABEL.test((b.textContent || '').trim()));
    (safe || buttons[buttons.length - 1])?.focus({ preventScroll: true });
  };
  const raf = requestAnimationFrame(placeFocus);

  // ⑤ 입력 추적 — 사용자가 친 것만 센다(화면이 값을 채워 넣는 것은 이 이벤트가 나지 않는다)
  const onInput = (e: Event) => {
    const t = e.target as Element | null;
    if (t && root.contains(t) && !isScratch(t)) dirty = true;
  };
  root.addEventListener('input', onInput, true);
  root.addEventListener('change', onInput, true);

  // ③④ 키 — document 에서 받는다(화면 부품의 자체 처리가 먼저 끝난 뒤라 가로채지 않는다)
  const onKey = (e: KeyboardEvent) => {
    if (!opts.isTop() || e.isComposing) return;
    const t = e.target as HTMLElement | null;
    if (!t || !root.contains(t)) return;

    if (e.key === 'Tab' && opts.trapTab) {
      const items = list(root, FOCUSABLE);
      if (!items.length) return;
      const first = items[0], last = items[items.length - 1];
      if (e.shiftKey && document.activeElement === first) { e.preventDefault(); last.focus(); }
      else if (!e.shiftKey && document.activeElement === last) { e.preventDefault(); first.focus(); }
      return;
    }
    if (e.key !== 'Enter' || e.defaultPrevented) return;

    if (e.ctrlKey || e.metaKey) {
      const btn = findPrimaryButton(root);
      if (btn) { e.preventDefault(); btn.click(); }
      return;
    }
    // 한 줄 칸에서 Enter = 다음 칸. 여러 줄 칸·버튼·자체 처리 칸은 그대로 둔다
    if (e.shiftKey || e.altKey) return;
    if (!(t instanceof HTMLInputElement)) return;
    if (['checkbox', 'radio', 'file', 'button', 'submit', 'search'].includes(t.type)) return;
    if (isScratch(t) || hasOwnKeyHandler(t)) return;
    const fields = list(root, FIELD);
    const next = fields[fields.indexOf(t) + 1];
    e.preventDefault();
    if (next) next.focus();
    else findPrimaryButton(root)?.focus();
  };
  document.addEventListener('keydown', onKey);

  return {
    isDirty: () => dirty,
    dispose: () => {
      cancelAnimationFrame(raf);
      root.removeEventListener('input', onInput, true);
      root.removeEventListener('change', onInput, true);
      document.removeEventListener('keydown', onKey);
      // ② 연 버튼으로 — 그 버튼이 아직 화면에 있을 때만
      if (opener && document.contains(opener)) requestAnimationFrame(() => { if (document.contains(opener)) opener.focus({ preventScroll: true }); });
    },
  };
}
