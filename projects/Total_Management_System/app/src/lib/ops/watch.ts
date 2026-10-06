/**
 * 운영 자동 점검 (ops-watch) — 2026-10-06
 *
 * "멈췄는데 아무도 모르는" 상태를 잡는다. Claude/외부 AI 없이 TMS 크론만으로 동작 (추가 비용 0).
 *   10분마다 : Make 시나리오 꺼짐/오류 상태  ← 꺼지면 알림톡 전부 정지 (2026-09-17 실사고)
 *   1시간마다: Make 미완료 실행(DLQ) 증가 · 실행 오류(로그 status 3)
 *   매일 08시: 연결·안전망 상태 + 복원수리 멈춘 건
 *
 * 알림 원칙 (알림 피로 방지)
 *   · 이상이 있을 때만. "정상" 알림 없음
 *   · 같은 문제는 한 번만 — 상태(system_settings 'ops.watch_state')에 보고한 키를 기억
 *     해결되면 키가 빠지고, 다시 생기면 다시 알린다
 *   · 매일 점검은 "새로 생긴 항목"이 있을 때만 발송(기존 항목은 목록에 같이 보여줌)
 *   · 발송 = sendPushToAll → 휴대폰이 2분 안에 못 받으면 메일 자동 예비발송(push-fallback)
 *
 * 송장 할 일 권한 문제는 shipping-todo 크론이 직접 알린다 → 여기서 중복으로 다루지 않음.
 */

import { sendPushToAll } from '@/lib/firebase/send-push';
import { isMobileLabel } from '@/lib/firebase/push-fallback';
import { getOwnerAlertStatus } from '@/lib/notification/owner-alert';
import { REPAIR_STATUS_LABEL } from '@/lib/repair/transitions';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Db = any;

const STATE_KEY = 'ops.watch_state';
const DAILY_HOUR_KST = 8;
const MAKE_BASE = (process.env.MAKE_API_BASE || 'https://eu2.make.com/api/v2').replace(/\/$/, '');
const MAKE_TEAM_ID = process.env.MAKE_TEAM_ID || '1942714';

interface WatchState {
  makeDown: string[];                  // 보고한 "꺼짐/오류" 시나리오 id
  dlq: Record<string, number>;         // 마지막으로 본 미완료 실행 수
  lastHourly: string;                  // 'YYYY-MM-DDTHH' (KST)
  lastHourlyAt: string;                // ISO — 오류 로그 조회 기준 시각
  lastDaily: string;                   // 'YYYY-MM-DD' (KST)
  daily: string[];                     // 마지막 매일 점검에서 보고된 항목 키
}

export interface WatchResult {
  ran: string[];                       // 'make' | 'hourly' | 'daily'
  alerts: Array<{ title: string; body: string }>;
  daily?: Array<{ key: string; text: string; isNew: boolean }>;
  makeConfigured: boolean;
  dryRun: boolean;
  errors: string[];
}

function kstParts(d = new Date()) {
  const k = new Date(d.getTime() + 9 * 3600 * 1000);
  const date = k.toISOString().slice(0, 10);
  return { date, hour: k.getUTCHours(), hourKey: `${date}T${String(k.getUTCHours()).padStart(2, '0')}` };
}
const daysSince = (iso: string | null | undefined) =>
  iso ? (Date.now() - new Date(iso).getTime()) / 86400000 : 0;
const mmdd = (d: string) => `${Number(d.slice(5, 7))}/${Number(d.slice(8, 10))}`;

async function loadState(db: Db): Promise<WatchState> {
  const empty: WatchState = { makeDown: [], dlq: {}, lastHourly: '', lastHourlyAt: '', lastDaily: '', daily: [] };
  try {
    const { data } = await db.from('system_settings').select('value').eq('key', STATE_KEY).maybeSingle();
    const v = typeof data?.value === 'string' ? JSON.parse(data.value) : data?.value;
    return v && typeof v === 'object' ? { ...empty, ...v } : empty;
  } catch { return empty; }
}
async function saveState(db: Db, s: WatchState) {
  const now = new Date().toISOString();
  await db.from('system_settings').upsert({ key: STATE_KEY, value: JSON.stringify(s), updated_at: now }, { onConflict: 'key' });
}

// ── Make ───────────────────────────────────────────────
interface MakeScenario { id: number; name: string; isActive: boolean; isinvalid?: boolean; dlqCount?: number }

async function makeGet<T>(path: string): Promise<T> {
  const token = (process.env.MAKE_API_TOKEN || '').trim();
  const res = await fetch(`${MAKE_BASE}${path}`, { headers: { Authorization: `Token ${token}` }, cache: 'no-store' });
  if (!res.ok) throw new Error(`Make API ${res.status} ${path}`);
  return res.json() as Promise<T>;
}
async function listScenarios(): Promise<MakeScenario[]> {
  const j = await makeGet<{ scenarios: MakeScenario[] }>(`/scenarios?teamId=${MAKE_TEAM_ID}&pg%5Blimit%5D=100`);
  return j.scenarios || [];
}

// ── 매일 점검 항목 ─────────────────────────────────────
async function collectDaily(db: Db, scenarios: MakeScenario[] | null, today: string): Promise<Array<{ key: string; text: string }>> {
  const items: Array<{ key: string; text: string }> = [];

  // 1) Make: 토큰 / "미완료 실행 저장" 꺼진 시나리오 (꺼져 있으면 오류 1번에 시나리오 전체가 멈춘다)
  if (!process.env.MAKE_API_TOKEN) {
    items.push({ key: 'make:no-token', text: 'Make 점검 불가 — Vercel 환경변수 MAKE_API_TOKEN 미설정' });
  } else if (!scenarios) {
    items.push({ key: 'make:api-fail', text: 'Make 상태 조회 실패 — 토큰 만료 여부 확인 (MAKE_API_TOKEN)' });
  } else {
    for (const sc of scenarios) {
      try {
        const bp = await makeGet<{ response?: { blueprint?: { metadata?: { scenario?: { dlq?: boolean } } } } }>(`/scenarios/${sc.id}/blueprint`);
        if (bp.response?.blueprint?.metadata?.scenario?.dlq !== true) {
          items.push({ key: `make:dlq-off:${sc.id}`, text: `Make 「${sc.name}」 미완료 실행 저장(Store incomplete executions)이 꺼져 있음` });
        }
      } catch { /* 개별 조회 실패는 다음 날 재시도 */ }
    }
  }

  // 2) 예비 메일 경로 (푸시가 휴대폰에 안 갔을 때 쓰는 길) — 끊겨 있으면 안전망이 없는 상태
  const mail = await getOwnerAlertStatus();
  if (!mail.ready) items.push({ key: 'mail:not-ready', text: `예비 알림 메일 사용 불가 — ${(mail.problem || '').slice(0, 80)}`.trim() });

  // 3) 휴대폰 알림 등록
  const { data: subs } = await db.from('push_subscriptions').select('device_info');
  if (!(subs || []).some((s: { device_info: string | null }) => isMobileLabel(s.device_info))) {
    items.push({ key: 'push:no-mobile', text: '휴대폰 알림 미등록 — 휴대폰에서 TMS 앱을 열어 알림 허용' });
  }

  // 4) 아임웹 연결 (주문 동기화·재고 반영이 여기에 기댄다)
  //    읽기만 한다 — 점검이 토큰을 갱신(rotation)하면 안 됨. 12시간 keep-alive 크론이 갱신하므로 25시간 넘게 멈췄으면 끊긴 것
  const { data: tok } = await db.from('system_settings').select('value').eq('key', 'imweb_openapi.token_updated_at').maybeSingle();
  const tokAgeH = tok?.value ? (Date.now() - new Date(String(tok.value)).getTime()) / 3600000 : Infinity;
  if (tokAgeH > 25) {
    items.push({ key: 'imweb:token', text: `아임웹 연결 갱신이 ${Number.isFinite(tokAgeH) ? `${Math.floor(tokAgeH)}시간째` : ''} 멈춤 — 주문 동기화 확인, 필요 시 설정에서 아임웹 재연결`.replace('  ', ' ') });
  }

  // 5) 복원수리 — 단계가 멈춘 건 (상태가 바뀌면 키가 바뀌어 다시 판단)
  const { data: repairs } = await db.from('repairs')
    .select('id, as_id, name, status, proceed_type, visit_date, pickup_date, received_at, updated_at, shipped_at')
    .not('status', 'in', '(completed,cancelled,delivered)');
  for (const r of repairs || []) {
    const label = REPAIR_STATUS_LABEL[r.status] || r.status;
    const who = `${r.as_id} ${r.name || ''}`.trim();
    let why = '';
    if (r.status === 'intake' && r.proceed_type === '직접방문') {
      if (r.visit_date && r.visit_date < today) why = `방문일(${mmdd(r.visit_date)}) 지났는데 ${label}`;
    } else if (r.status === 'intake') {
      if (daysSince(r.received_at) >= 7) why = `접수 후 ${Math.floor(daysSince(r.received_at))}일째 ${label}`;
    } else if (r.status === 'pickup_scheduled') {
      const base = r.pickup_date ? `${r.pickup_date}T00:00:00+09:00` : r.updated_at;
      if (daysSince(base) >= 3) why = `${r.pickup_date ? `수거일(${mmdd(r.pickup_date)})` : '수거 예약'} 후 ${Math.floor(daysSince(base))}일째 ${label}`;
    } else if (r.status === 'shipped') {
      const base = r.shipped_at || r.updated_at;
      if (daysSince(base) >= 7) why = `출고 후 ${Math.floor(daysSince(base))}일째 배송완료 안 됨`;
    } else if (daysSince(r.updated_at) >= 3) {
      why = `${Math.floor(daysSince(r.updated_at))}일째 ${label}`;
    }
    if (why) items.push({ key: `repair:${r.id}:${r.status}`, text: `복원수리 ${who} — ${why}` });
  }

  return items;
}

export async function runOpsWatch(db: Db, opts: { dryRun?: boolean; force?: 'hourly' | 'daily' } = {}): Promise<WatchResult> {
  const dryRun = !!opts.dryRun;
  const state = await loadState(db);
  const now = kstParts();
  const result: WatchResult = { ran: [], alerts: [], makeConfigured: !!process.env.MAKE_API_TOKEN, dryRun, errors: [] };
  const alert = (title: string, body: string) => result.alerts.push({ title, body });

  // ── 매 실행: Make 시나리오 꺼짐 ──
  let scenarios: MakeScenario[] | null = null;
  if (result.makeConfigured) {
    try {
      scenarios = await listScenarios();
      result.ran.push('make');
      const down = scenarios.filter((s) => !s.isActive || s.isinvalid);
      const fresh = down.filter((s) => !state.makeDown.includes(String(s.id)));
      if (fresh.length) {
        alert('⚠️ Make 시나리오 꺼짐 — 알림톡 정지',
          `${fresh.map((s) => `「${s.name}」${s.isinvalid ? ' (설정 오류)' : ''}`).join(', ')} · Make에서 다시 켜고 「Process old data」로 밀린 건 처리`);
      }
      state.makeDown = down.map((s) => String(s.id));
    } catch (e) {
      result.errors.push(String(e));
    }
  }

  // ── 1시간마다: 미완료 실행 증가 · 실행 오류 ──
  if (scenarios && (opts.force === 'hourly' || state.lastHourly !== now.hourKey)) {
    result.ran.push('hourly');
    const since = state.lastHourlyAt || new Date(Date.now() - 3600 * 1000).toISOString();
    const lines: string[] = [];
    for (const sc of scenarios) {
      const id = String(sc.id);
      const cnt = sc.dlqCount || 0;
      if (cnt > (state.dlq[id] || 0)) lines.push(`「${sc.name}」 미완료 실행 ${cnt}건`);
      state.dlq[id] = cnt;
      try {
        const logs = await makeGet<{ scenarioLogs?: Array<{ status: number; timestamp: string; eventType?: string }> }>(
          `/scenarios/${sc.id}/logs?pg%5Blimit%5D=50`);
        // status 1=성공 2=경고 3=오류. "미완료 실행 저장"이 켜져 있으면 오류가 경고(2)로 남는다 → 둘 다 잡음
        const errs = (logs.scenarioLogs || []).filter((l) => l.eventType === 'EXECUTION_END' && l.status >= 2 && l.timestamp > since);
        if (errs.length) lines.push(`「${sc.name}」 실행 오류 ${errs.length}건`);
      } catch (e) { result.errors.push(String(e)); }
    }
    if (lines.length) alert('⚠️ 알림톡 발송 오류 확인 필요', `${lines.join(' · ')} · Make 실행 기록에서 확인`);
    state.lastHourly = now.hourKey;
    state.lastHourlyAt = new Date().toISOString();
  }

  // ── 매일 08시: 연결 상태 + 멈춘 건 ──
  if (opts.force === 'daily' || (now.hour >= DAILY_HOUR_KST && state.lastDaily !== now.date)) {
    result.ran.push('daily');
    try {
      const items = await collectDaily(db, scenarios, now.date);
      const marked = items.map((i) => ({ ...i, isNew: !state.daily.includes(i.key) }));
      result.daily = marked;
      const newCount = marked.filter((i) => i.isNew).length;
      if (newCount) {
        const body = marked.map((i) => `${i.isNew ? '🆕 ' : '• '}${i.text}`).join('\n');
        alert(`아침 점검 — 확인할 것 ${marked.length}건 (새로 ${newCount}건)`, body);
      }
      state.daily = items.map((i) => i.key);
      state.lastDaily = now.date;
    } catch (e) {
      result.errors.push(String(e));
    }
  }

  if (!dryRun) {
    for (const a of result.alerts) {
      try {
        await sendPushToAll({ title: a.title, body: a.body, url: '/dashboard', tag: `mamoru-ops-${a.title.slice(0, 12)}` });
      } catch (e) { result.errors.push(`push: ${String(e)}`); }
    }
    await saveState(db, state);
  }
  return result;
}
