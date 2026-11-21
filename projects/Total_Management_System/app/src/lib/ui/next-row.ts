/**
 * 줄 끝 버튼으로 처리한 뒤 다음 건으로 (2026-11-21)
 *
 * 처리하면 그 줄이 탭에서 빠지거나 순서가 바뀐다. 그래서 "다음 줄"은 처리하기 전에 기억해 두고,
 * 처리가 끝난 뒤 그 줄을 눌러 준다(목록을 다시 불러와 줄이 새로 그려져도 찾도록 줄 열쇠로도 기억).
 * 목록 줄 표시(data-list-row · data-row-key)는 DataGrid 와 각 목록이 붙인다.
 */
export function rememberNextRow(from: Element | null): () => void {
  const row = from?.closest('[data-list-row]') ?? null;
  let next = row?.nextElementSibling ?? null;
  while (next && !next.hasAttribute('data-list-row')) next = next.nextElementSibling;
  const key = next?.getAttribute('data-row-key') ?? null;
  const el = next as HTMLElement | null;
  return () => {
    if (!el && !key) return; // 마지막 줄이었다
    window.setTimeout(() => {
      const byKey = key ? document.querySelector<HTMLElement>(`[data-list-row][data-row-key="${CSS.escape(key)}"]`) : null;
      const target = byKey ?? (el && el.isConnected ? el : null);
      if (!target) return;
      target.click();
      target.scrollIntoView({ block: 'nearest' });
    }, 400);
  };
}
