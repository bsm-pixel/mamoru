/**
 * Google Tasks API wrapper (2026-09-24)
 *
 * 왜 캘린더 '일정'이 아니라 '할 일'인가:
 *   송장 생성은 약속이 아니라 **처리해야 할 일**이다. 할 일로 넣어야 캘린더에서 체크(○→✓)로 끝낼 수 있고,
 *   구글 캘린더/할일 앱 어디서나 '남은 일'로 모인다. 일정으로 넣으면 체크가 안 돼 끝냈는지 알 수 없다.
 *
 * 권한: calendar.events 와 별개로 `auth/tasks` 스코프가 필요 → 설정에서 구글 재연결 1회 필요.
 * 실패는 전부 내부에서 삼키고 결과만 반환 — 판매/납품 로직을 막지 않는다.
 */

import { google } from 'googleapis';
import { getAuthorizedClient } from './oauth';

export interface TaskResult {
  ok: boolean;
  taskId?: string;
  error?: string;
  notConnected?: boolean;
  needsReauth?: boolean;   // 토큰에 tasks 스코프가 없음 → 재연결 안내
}

/** 권한 부족(스코프 미승인) 판별 — 재연결 안내를 위해 일반 오류와 구분한다 */
function isScopeError(msg: string): boolean {
  return /insufficient|ACCESS_TOKEN_SCOPE|invalid_scope|Request had insufficient authentication scopes|403/i.test(msg);
}

/** 할 일 생성 — due 는 날짜만 의미 있다(구글이 시각은 무시) */
export async function createTask(params: {
  title: string;
  notes?: string;
  due?: string;            // 'YYYY-MM-DD' (KST 기준 날짜)
  taskListId?: string;
}): Promise<TaskResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };

    const tasks = google.tasks({ version: 'v1', auth });
    const res = await tasks.tasks.insert({
      tasklist: params.taskListId || '@default',
      requestBody: {
        title: params.title,
        notes: params.notes,
        due: params.due ? `${params.due}T00:00:00.000Z` : undefined,
      },
    });
    if (!res.data.id) return { ok: false, error: 'no_task_id_returned' };
    return { ok: true, taskId: res.data.id };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, error: msg, needsReauth: isScopeError(msg) };
  }
}

/** 할 일 삭제 (이미 없으면 ok 로 본다 — 지워진 걸 또 지울 필요는 없다) */
export async function deleteTask(taskId: string, taskListId = '@default'): Promise<TaskResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };

    const tasks = google.tasks({ version: 'v1', auth });
    await tasks.tasks.delete({ tasklist: taskListId, task: taskId });
    return { ok: true };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (/\b(404|410|notFound|deleted)\b/i.test(msg)) return { ok: true };
    return { ok: false, error: msg, needsReauth: isScopeError(msg) };
  }
}

/** 연결 점검용 — 기본 할 일 목록 이름을 읽어본다 */
export async function probeTasks(): Promise<TaskResult & { listTitle?: string }> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };
    const tasks = google.tasks({ version: 'v1', auth });
    const res = await tasks.tasklists.get({ tasklist: '@default' });
    return { ok: true, listTitle: res.data.title || '' };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return { ok: false, error: msg, needsReauth: isScopeError(msg) };
  }
}
