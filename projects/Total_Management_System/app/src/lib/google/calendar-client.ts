/**
 * Google Calendar API wrapper
 * insert / patch / delete 3종 + 최근 오류 로깅
 */

import { google, calendar_v3 } from 'googleapis';
import { getAuthorizedClient } from './oauth';
import { createServiceClient } from '@/lib/supabase/server';

export interface CalendarResult {
  ok: boolean;
  eventId?: string;
  error?: string;
  notConnected?: boolean;
}

/** 이벤트 생성 */
export async function createCalendarEvent(params: {
  calendarId?: string;
  event: calendar_v3.Schema$Event;
}): Promise<CalendarResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };

    const calendar = google.calendar({ version: 'v3', auth });
    const res = await calendar.events.insert({
      calendarId: params.calendarId || 'primary',
      requestBody: params.event,
    });

    if (!res.data.id) return { ok: false, error: 'no_event_id_returned' };

    await logSuccess();
    return { ok: true, eventId: res.data.id };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    await logLastError(msg);
    return { ok: false, error: msg };
  }
}

/** 이벤트 부분 업데이트 (PATCH) */
export async function updateCalendarEvent(params: {
  eventId: string;
  calendarId?: string;
  event: calendar_v3.Schema$Event;
}): Promise<CalendarResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };

    const calendar = google.calendar({ version: 'v3', auth });
    await calendar.events.patch({
      calendarId: params.calendarId || 'primary',
      eventId: params.eventId,
      requestBody: params.event,
    });

    await logSuccess();
    return { ok: true, eventId: params.eventId };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    await logLastError(msg);
    return { ok: false, error: msg };
  }
}

/** 이벤트 삭제 (이미 없어진 경우 성공 취급) */
export async function deleteCalendarEvent(params: {
  eventId: string;
  calendarId?: string;
}): Promise<CalendarResult> {
  try {
    const auth = await getAuthorizedClient();
    if (!auth) return { ok: false, notConnected: true, error: 'not_connected' };

    const calendar = google.calendar({ version: 'v3', auth });
    await calendar.events.delete({
      calendarId: params.calendarId || 'primary',
      eventId: params.eventId,
    });

    await logSuccess();
    return { ok: true };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    // 이미 삭제된 경우 (404 / 410) → 성공 취급
    if (/\b(404|410|notFound|deleted)\b/i.test(msg)) {
      return { ok: true };
    }
    await logLastError(msg);
    return { ok: false, error: msg };
  }
}

async function logLastError(msg: string): Promise<void> {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;
    await dbAny.from('system_settings').upsert(
      {
        key: 'google.calendar.last_error',
        value: `${new Date().toISOString()} :: ${msg.slice(0, 500)}`,
      },
      { onConflict: 'key' },
    );
  } catch {
    /* 로깅 실패는 무시 */
  }
}

async function logSuccess(): Promise<void> {
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;
    await dbAny.from('system_settings').upsert(
      {
        key: 'google.calendar.last_success_at',
        value: new Date().toISOString(),
      },
      { onConflict: 'key' },
    );
    // 성공 시 이전 에러는 지움
    await dbAny
      .from('system_settings')
      .delete()
      .eq('key', 'google.calendar.last_error');
  } catch {
    /* 로깅 실패는 무시 */
  }
}
