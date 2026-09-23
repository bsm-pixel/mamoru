/**
 * 송장 미생성 건 → Google Calendar '할 일' 종일 일정 (2026-09-24 신규)
 *
 * 상담(calendar-sync.ts)·복원수리(repair-calendar-sync.ts)와 같은 패턴.
 * 다른 점: 고객 약속이 아니라 **사장님 업무 할 일**이라 시간이 없다 → 종일 일정.
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
 * 이벤트 id 보관: system_settings `calendar.shipping_todo_events` = { "sale:<id>": "<eventId>" }
 *   컬럼을 안 쓴 이유 — 마이그레이션 없이 시작하기 위함. 쓰는 쪽이 이 크론 하나뿐이라 경합 없음.
 */

import { createServiceClient } from '@/lib/supabase/server';
import { createCalendarEvent, deleteCalendarEvent } from './calendar-client';
import { formatShippingTodoToEvent } from './event-formatter';

const BASE_URL = process.env.NEXT_PUBLIC_APP_URL || 'https://app-eta-sandy-75.vercel.app';
const MAP_KEY = 'calendar.shipping_todo_events';
const SINCE_KEY = 'calendar.shipping_todo_since';
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
      .in('key', [MAP_KEY, SINCE_KEY]);
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
    const eventMap = parse(raw[MAP_KEY]);
    const since = (raw[SINCE_KEY] || SINCE_DEFAULT).replace(/^"|"$/g, '').slice(0, 10);
    const today = /^\d{4}-\d{2}-\d{2}$/.test(opts.asOf || '') ? (opts.asOf as string) : kstDate();
    const tomorrow = kstDate(new Date(new Date(`${today}T00:00:00+09:00`).getTime() + 24 * 60 * 60 * 1000));

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

    // 3) 정리 — 더는 대기가 아닌 건(송장 생성·출고·취소·완료)의 일정 삭제
    for (const [key, eventId] of Object.entries(eventMap)) {
      if (pendingKeys.has(key)) continue;
      if (!opts.dryRun) {
        const res = await deleteCalendarEvent({ eventId });
        // 404/410(이미 지워짐)도 맵에서는 제거 — 남겨두면 영영 안 지워진다
        if (!res.ok && !/\b(404|410|notFound|deleted)\b/i.test(res.error || '')) continue;
        delete eventMap[key];
        mapChanged = true;
      }
      result.deleted.push(key);
    }

    // 4) 생성 — 등록 '다음 날'부터, 아직 일정이 없는 건만
    for (const p of pendings) {
      if (eventMap[p.key]) continue;
      if (p.createdDate >= today) { result.pending.push(p.key); continue; }   // 당일 등록분은 내일 아침에
      if (!opts.create) continue;
      if (!opts.dryRun) {
        const res = await createCalendarEvent({
          // 제목·본문 규칙은 상담/수리와 같은 곳(event-formatter)에서 관리 — 형식이 갈라지지 않게
          event: formatShippingTodoToEvent(
            { kind: p.kind, who: p.who, docNo: p.docNo, amount: p.amount, createdAt: p.createdAt, date: today, nextDate: tomorrow },
            BASE_URL,
          ),
        });
        if (!res.ok || !res.eventId) continue;   // 미연결·오류는 조용히 — 다음 실행에서 재시도
        eventMap[p.key] = res.eventId;
        mapChanged = true;
      }
      result.created.push(p.key);
    }

    // 5) 맵 저장
    if (mapChanged && !opts.dryRun) {
      await dbAny.from('system_settings').upsert(
        { key: MAP_KEY, value: JSON.stringify(eventMap), updated_at: new Date().toISOString() },
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
