/**
 * 저장 버튼이 흐릴 때 "무엇이 빠졌는지" 한 줄 (2026-11-15)
 * 전엔 버튼만 흐려져 있어 왜 안 눌리는지 알 수 없었다. 빠진 것이 없으면 아무것도 그리지 않는다.
 *   <MissingNote items={[!name.trim() && '이름', !amount && '금액']} />
 */
export function MissingNote({ items, className = '' }: { items: Array<string | false | null | undefined>; className?: string }) {
  const list = items.filter(Boolean) as string[];
  if (list.length === 0) return null;
  return (
    <p role="status" className={`text-[11px] text-neutral-500 ${className}`}>
      빠진 항목: <span className="font-semibold text-neutral-700">{list.join(' · ')}</span>
    </p>
  );
}
