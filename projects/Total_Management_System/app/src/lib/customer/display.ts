/**
 * 고객 표시명 헬퍼 — 활동명(매장 사용 이름) 있으면 우선 표시.
 * 상담·출장 화면 위주 사용. 판매·계약 등 실명이 중요한 곳에는 적용하지 않음.
 */

function clean(s?: string | null): string {
  return (s ?? '').trim();
}

/** 카드/상세 표시명 — 활동명 있으면 "하은 (김순실)", 없으면 "김순실" */
export function activityDisplay(activityName?: string | null, name?: string | null): string {
  const a = clean(activityName);
  const n = clean(name);
  if (a && a !== n) return n ? `${a} (${n})` : a;
  return n;
}

/** 호칭용(~님) — 활동명 있으면 활동명, 없으면 실명 */
export function honorific(activityName?: string | null, name?: string | null): string {
  return clean(activityName) || clean(name);
}

/** 직급까지 — "하은 원장" (직급 있을 때), 없으면 활동명/실명 */
export function activityWithPosition(activityName?: string | null, name?: string | null, position?: string | null): string {
  const base = honorific(activityName, name);
  const p = clean(position);
  return p ? `${base} ${p}` : base;
}

/** 실명 옆 보조표기 — "박봄 디자이너" (활동명·직급). 없으면 ''. 고객 화면에서 실명 뒤 작게/흐리게 표시용 */
export function activitySuffix(activityName?: string | null, position?: string | null): string {
  return [clean(activityName), clean(position)].filter(Boolean).join(' ');
}
