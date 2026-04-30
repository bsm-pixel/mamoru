/**
 * 에러를 사람이 읽을 수 있는 string으로 직렬화 (서버/클라이언트 공용)
 *
 * `[object Object]` 표시 방지 — Supabase/외부 SDK가 던지는 객체 에러를
 * 토스트/응답에 안전하게 노출.
 */
export function errMsg(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (typeof err === 'string') return err;
  if (err && typeof err === 'object') {
    const e = err as { message?: unknown; error?: unknown; details?: unknown; hint?: unknown };
    if (typeof e.message === 'string') return e.message;
    if (typeof e.error === 'string') return e.error;
    if (typeof e.details === 'string') return e.details;
    if (typeof e.hint === 'string') return e.hint;
    try {
      return JSON.stringify(err);
    } catch {
      return String(err);
    }
  }
  return String(err);
}
