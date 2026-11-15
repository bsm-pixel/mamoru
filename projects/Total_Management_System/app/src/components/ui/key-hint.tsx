/**
 * 버튼 옆의 작은 단축키 글씨 (2026-11-15) — 모바일에서는 숨긴다
 *   <Button>판매 입력 <KeyHint>F2</KeyHint></Button>
 */
export function KeyHint({ children, className = '' }: { children: string; className?: string }) {
  return (
    <kbd aria-hidden className={`hidden lg:inline-block ml-1 px-1 py-px rounded border border-current/25 text-[10px] font-sans font-medium leading-none opacity-60 ${className}`}>
      {children}
    </kbd>
  );
}
