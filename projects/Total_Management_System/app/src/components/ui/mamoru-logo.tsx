import { MAMORU_WORDMARK_ASPECT, MAMORU_WORDMARK_PATH, MAMORU_WORDMARK_VIEWBOX } from '@/lib/brand/logo-svg';

/**
 * MAMORU 워드마크(로고 글씨) — AGRESSIVE 서체. "MAMORU"를 로고로 쓸 때는 글자 대신 이 컴포넌트.
 * 크기는 width/height 속성으로 고정 → Tailwind 가 없는 인쇄용 새 창(innerHTML 복사)에서도 그대로 나온다.
 * 색은 기본 currentColor(부모 글자색). html2canvas 로 캡처되는 영역은 color 를 직접 지정한다.
 */
export function MamoruWordmark({ height = 14, color = 'currentColor', className, style }: {
  height?: number;
  color?: string;
  className?: string;
  style?: React.CSSProperties;
}) {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={MAMORU_WORDMARK_VIEWBOX}
      width={Math.round(height * MAMORU_WORDMARK_ASPECT * 100) / 100}
      height={height}
      role="img"
      aria-label="MAMORU"
      className={className}
      style={{ display: 'inline-block', verticalAlign: 'middle', ...style }}
    >
      <path fill={color} d={MAMORU_WORDMARK_PATH} />
    </svg>
  );
}
