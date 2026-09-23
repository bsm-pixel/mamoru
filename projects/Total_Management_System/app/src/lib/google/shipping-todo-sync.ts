/**
 * 송장 미생성 건 → Google **Tasks(할 일)** (2026-09-24)
 *
 * 상담(calendar-sync.ts)·복원수리(repair-calendar-sync.ts)와 같은 패턴.
 * 고객 약속이 아니라 **처리해야 할 일**이다 → 일정이 아니라 할 일(체크로 끝낼 수 있음).
 * 캘린더 '종일 일정'으로 넣던 최초 구현은 체크가 안 돼 폐기(09-24). 남은 일정은 자동 정리한다.
 *
 * 대상 (실데이터로 기준 확정 — 2026-09-24 조회)
 *   · 판매 offline_sales : delivery_method='shipping' AND invoice_number IS NULL
 *       - 'pickup'(매장 수령)은 원래 송장이 없다 → 제외. 실측 shipping 66건은 전부 송장 있음(오탐 0)
 *   · 납품 deliveries    : status='confirmed' AND tracking_number IS NULL
 *       - draft(작성중)·shipped(직접전달 포함)·cancelled 제외
 *
 * 규칙
 *   · 등록 **다음 날**부터 올린다 (당일 처리 흐름을 방해하지 않음)
 *   · since(=calendar.shipping_todo_since) 이후 등록분만 — 옛 데이터 소급 금지
 *   · 송장 생성·출고·취소되면 다음 정리 때 일정이 사라진다 (경로마다 손대지 않고 여기서 일괄)
 *
 * 할 일 id 보관: system_settings `calendar.shipping_todo_tasks` = { "sale:<id>": "<taskId>" }
 *   컬럼을 안 쓴 이유 — 마이그레이션 없이 시작하기 위함. 쓰는 쪽이 이 크론 하나뿐이라 경합 없음.
 */

import { createServiceClient } from '@/lib/supabase/server';
import { deleteCalendarEvent } from './calendar-client';
import { createTask, deleteTask } from './tasks-client';
import { formatShippingTodoToTask } from './event-formatter';

const MAP_KEY = 'calendar.shipping_todo_tasks';              // { "sale:<id>": "<taskId>" }
const LEGACY_EVENT_MAP_KEY = 'calendar.shipping_todo_events'; // 09-24 이전: 캘린더 '일정'으로 넣던 시절 — 정리 후 비움
const SINCE_KEY = 'calendar.shipping_todo_since';
const LAST_RUN_KEY = 'calendar.shipping_todo_last_run';   // 크론이 실제로 돌고 있는지 확인용(무소식이면 크론 미등록)
const SINCE_DEFAULT = '2026-09-24';   // 기능 시작일 — 이전 등록분은 올리지 않는다

/** KST 기준 YYYY-MM-DD */
export function kstDate(d: Date = new Date()): string {
  return new Date(d.getTime() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
}

/** KST 기준 0~23시 */
export function kstHour(d: Date = new Date()): number {
  return new Date(d.getTime() + 9 * 60 * 60 * 1000).getUTCHours();
}

interface Pending {
  key: string;          // sale:<id> | delivery:<id>
  kind: 'sale' | 'delivery';
  who: string;
  docNo: string | null;
  amount: number | null;
  createdAt: string | null;
  createdDate: string;  // KST 등록일
}

export interface SweepResult {
  created: string[];
  deleted: string[];
  pending: string[];    // 대상이지만 아직 '다음 날'이 안 된 건
  dryRun: boolean;
  /** 토큰에 tasks 스코프가 없음 → TMS 설정에서 구글 재연결 필요 */
  needsReauth?: boolean;
  error?: string;
}

/**
 * 송장 미생성 건 훑기.
 * @param opts.create false 면 정리(삭제)만 한다 — 생성은 하루 1회(아침)만 하기 위함
 * @param opts.asOf   'YYYY-MM-DD' 기준일 대체 — 검증·소급 처리용(미지정=오늘 KST)
 */
export async function sweepShippingTodos(opts: { create: boolean; dryRun?: boolean; asOf?: string }): Promise<SweepResult> {
  const result: SweepResult = { created: [], deleted: [], pending: [], dryRun: !!opts.dryRun };
  try {
    const db = createServiceClient();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const dbAny = db as any;

    // 1) 설정 로드 (이벤트 맵 + 시작일)
    const { data: settingRows } = await dbAny
      .from('system_settings')
      .select('key, value')
      .in('key', [MAP_KEY, SINCE_KEY, LEGACY_EVENT_MAP_KEY]);
    const raw: Record<string, string> = {};
    (settingRows || []).forEach((r: { key: string; value: string | null }) => {
      if (r.value != null) raw[r.key] = String(r.value);
    });
    const parse = (v: string | undefined): Record<string, string> => {
      if (!v) return {};
      try {
        const o = typeof v === 'string' ? JSON.parse(v) : v;
        return o && typeof o === 'object' && !Array.isArray(o) ? o : {};
      } catch { return {}; }
    };
    const taskMap = parse(raw[MAP_KEY]);
    const legacyEvents = parse(raw[LEGACY_EVENT_MAP_KEY]);
    const since = (raw[SINCE_KEY] || SINCE_DEFAULT).replace(/^"|"$/g, '').slice(0, 10);
    const today = /^\d{4}-\d{2}-\d{2}$/.test(opts.asOf || '') ? (opts.asOf as string) : kstDate();

    // 2) 지금 '송장 대기'인 건 수집
    const pendings: Pending[] = [];

    const { data: sales } = await dbAny
      .from('offline_sales')
      .select('id, sale_number, customer_name, total_amount, created_at, delivery_method, invoice_number, cancelled_at')
      .eq('delivery_method', 'shipping')
      .is('invoice_number', null)
      .is('cancelled_at', null)
      .gte('created_at', `${since}T00:00:00+09:00`);
    for (const s of sales || []) {
      pendings.push({
        key: `sale:${s.id}`,
        kind: 'sale',
        who: s.customer_name || '고객',
        docNo: s.sale_number || null,
        amount: s.total_amount ?? null,
        createdAt: s.created_at || null,
        createdDate: kstDate(new Date(s.created_at)),
      });
    }

    const { data: dels } = await dbAny
      .from('deliveries')
      .select('id, dl_number, customer_name, total_amount, created_at, status, tracking_number, cancelled_at')
      .eq('status', 'confirmed')
      .is('tracking_number', null)
      .is('cancelled_at', null)
      .gte('created_at', `${since}T00:00:00+09:00`);
    for (const d of dels || []) {
      pendings.push({
        key: `delivery:${d.id}`,
        kind: 'delivery',
        who: d.customer_name || '거래처',
        docNo: d.dl_number || null,
        amount: d.total_amount ?? null,
        createdAt: d.created_at || null,
        createdDate: kstDate(new Date(d.created_at)),
      });
    }

    const pendingKeys = new Set(pendings.map((p) => p.key));
    let mapChanged = false;

    // 3) 정리 — 더는 대기가 아닌 건(송장 생성·출고·취소·완료)의 할 일 삭제
    for (const [key, taskId] of Object.entries(taskMap)) {
      if (pendingKeys.has(key)) continue;
      if (!opts.dryRun) {
        const res = await deleteTask(String(taskId));   // 이미 지워졌으면 ok 로 돌아온다
        if (!res.ok) continue;
        delete taskMap[key];
        mapChanged = true;
      }
      result.deleted.push(key);
    }

    // 3-b) 옛 방식(캘린더 종일 일정)으로 남은 것 정리 — 같은 건이 '일정'과 '할 일' 둘로 보이지 않게 (09-24 전환)
    if (!opts.dryRun && Object.keys(legacyEvents).length > 0) {
      let legacyChanged = false;
      for (const [key, eventId] of Object.entries(legacyEvents)) {
        const res = await deleteCalendarEvent({ eventId: String(eventId) });
        if (!res.ok && !/\b(404|410|notFound|deleted)\b/i.test(res.error || '')) continue;
        delete legacyEvents[key];
        legacyChanged = true;
        result.deleted.push(`${key}(옛 일정)`);
      }
      if (legacyChanged) {
        await dbAny.from('system_settings').upsert(
          { key: LEGACY_EVENT_MAP_KEY, value: JSON.stringify(legacyEvents), updated_at: new Date().toISOString() },
          { onConflict: 'key' },
        );
      }
    }

    // 4) 생성 — 등록 '다음 날'부터, 아직 할 일이 없는 건만
    for (const p of pendings) {
      if (taskMap[p.key]) continue;
      if (p.createdDate >= today) { result.pending.push(p.key); continue; }   // 당일 등록분은 내일 아침에
      if (!opts.create) continue;
      if (!opts.dryRun) {
        const { title, notes } = formatShippingTodoToTask(p);
        const res = await createTask({ title, notes, due: today });
        if (!res.ok || !res.taskId) {
          // 스코프 미승인이면 조용히 실패해선 안 된다 — 설정에서 구글 재연결이 필요하다는 뜻
          if (res.needsReauth) result.needsReauth = true;
          continue;
        }
        taskMap[p.key] = res.taskId;
        mapChanged = true;
      }
      result.created.push(p.key);
    }

    // 5) 실행 흔적 — 할 일이 0건이면 아무 것도 안 바뀌어서 '크론이 도는지' 알 방법이 없다
    if (!opts.dryRun) {
      await dbAny.from('system_settings').upsert(
        { key: LAST_RUN_KEY, value: `${new Date().toISOString()} created=${result.created.length} deleted=${result.deleted.length} create=${opts.create}`, updated_at: new Date().toISOString() },
        { onConflict: 'key' },
      );
    }

    // 6) 맵 저장
    if (mapChanged && !opts.dryRun) {
      await dbAny.from('system_settings').upsert(
        { key: MAP_KEY, value: JSON.stringify(taskMap), updated_at: new Date().toISOString() },
        { onConflict: 'key' },
      );
    }
    return result;
  } catch (err) {
    result.error = err instanceof Error ? err.message : String(err);
    console.error('[shipping-todo] 훑기 실패', result.error);
    return result;
  }
}
