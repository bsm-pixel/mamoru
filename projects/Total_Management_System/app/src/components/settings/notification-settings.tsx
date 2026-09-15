'use client';

import { useState, useEffect } from 'react';
import { Button } from '@/components/ui/button';
import { Save, Bell, Send, Smartphone } from 'lucide-react';
import toast from 'react-hot-toast';
import type { TabProps } from '@/app/(dashboard)/settings/page';
import GoogleCalendarSettings from '@/components/settings/google-calendar-settings';
import BannerSettings from '@/components/settings/banner-settings';

function parse<T>(raw: unknown, fb: T): T {
  if (raw === undefined || raw === null) return fb;
  if (typeof raw === 'string') { try { return JSON.parse(raw); } catch { return raw as unknown as T; } }
  return raw as T;
}

export default function NotificationSettings({ settings, onSave, saving }: TabProps) {
  const [webhookConsultation, setWebhookConsultation] = useState('');
  const [webhookAsReceived, setWebhookAsReceived] = useState('');
  const [webhookRepair, setWebhookRepair] = useState('');
  const [webhookEvent, setWebhookEvent] = useState('');
  const [webhookImweb, setWebhookImweb] = useState(''); // 아임웹 주문 취소·반품 (2026-09-14) — 입력 = 가동
  // 🔴 앱 푸시 on/off 토글 제거(2026-08-01) — 고객 행동 푸시는 항상 발송(무조건). 놓치면 안 되므로 게이팅 없음.
  // 🔴 고객 알림톡 on/off 토글 제거(2026-09-14) — 항상 발송 원칙. 비상 정지는 해당 Make 웹훅 URL 비우기
  // 앱 화면 열려 있을 때 in-app 알림음(notification.wav) — 배송설정 탭에서 여기로 이동(2026-08-01), 기본 ON
  const [soundEnabled, setSoundEnabled] = useState(true);

  useEffect(() => {
    setWebhookConsultation(parse(settings['notifications.webhook_consultation'], ''));
    setWebhookAsReceived(parse(settings['notifications.webhook_as_received'], ''));
    setWebhookRepair(parse(settings['notifications.webhook_repair'], ''));
    setWebhookEvent(parse(settings['notifications.webhook_event'], ''));
    setWebhookImweb(parse(settings['notifications.webhook_imweb'], ''));
    setSoundEnabled(parse(settings['notifications.sound_enabled'], true));
  }, [settings]);

  const handleSave = () => {
    onSave([
      { key: 'notifications.webhook_consultation', value: webhookConsultation },
      { key: 'notifications.webhook_as_received', value: webhookAsReceived },
      { key: 'notifications.webhook_repair', value: webhookRepair },
      { key: 'notifications.webhook_event', value: webhookEvent },
      { key: 'notifications.webhook_imweb', value: webhookImweb },
      { key: 'notifications.sound_enabled', value: soundEnabled },
    ]);
  };

  return (
    <div className="space-y-6">
      <h2 className="text-lg font-bold">알림·연동 설정</h2>

      {/* ── Google Calendar 연동 ── */}
      <GoogleCalendarSettings />

      {/* ── 아임웹 배너/팝업 관리 ── */}
      <BannerSettings />

      {/* ── 내 푸시 알림 (사장님이 받는 것) ── */}
      <div className="rounded-lg border border-neutral-200 bg-stone-50 p-4 space-y-3">
        <div>
          <h3 className="text-sm font-bold text-stone-900">📱 내 푸시 알림</h3>
          <p className="text-xs text-neutral-500 mt-0.5">
            고객 접수·행동이 발생하면 사장님 디바이스로 <b>항상</b> 푸시 알림을 보냅니다. (크롬/모바일 앱)
          </p>
          <p className="text-[11px] text-neutral-400 mt-1">
            ※ 놓치면 안 되는 알림이라 on/off 설정 없이 무조건 발송됩니다. 안 오면 아래 <b>테스트</b>로 기기 연결을 확인하세요.
          </p>
        </div>

        {/* 앱 열려 있을 때 알림음 */}
        <div className="flex items-center justify-between py-1 border-t border-neutral-200 pt-3">
          <div>
            <span className="text-sm font-medium">앱 화면 열어둘 때 알림음 🔊</span>
            <p className="text-[11px] text-neutral-400 mt-0.5">
              TMS를 보고 있을 때 접수가 들어오면 &ldquo;띵&rdquo; 소리. (백그라운드/잠금 상태 소리는 휴대폰 알림 소리 설정을 따름)
            </p>
          </div>
          <Toggle checked={soundEnabled} onChange={setSoundEnabled} />
        </div>

        {/* 테스트 발송 패널 */}
        <PushTestPanel />

        {/* 디바이스 정리 — 중복 알림 해결 */}
        <DevicesPanel />
      </div>

      <div className="pt-4 border-t border-neutral-100">
        <h3 className="text-sm font-bold text-neutral-700 mb-3">💬 고객 알림톡 발송</h3>
      </div>

      {/* 항상 발송 원칙 (2026-09-14) — on/off 토글 없음 */}
      <div className="rounded-lg border border-neutral-200 bg-stone-50 p-3 text-xs text-neutral-600 leading-relaxed">
        고객 알림톡은 <b>항상 발송</b>됩니다 — 접수·확정·변경·출고·후기 요청 등 고객 행동과 상태 변화마다 빠짐없이 안내합니다.
        같은 내용이 두 번 가는 경우만 시스템이 자동으로 1통으로 줄입니다. (예: 무상 수리의 입금 확인, 매장 현장 결제)
        <br />
        <span className="text-neutral-400">※ 장애 등으로 특정 흐름을 잠시 멈춰야 하면 아래 해당 Make 웹훅 URL 을 비우고 저장하세요.</span>
      </div>

      {/* Make 웹훅 URL — 5개 시나리오 */}
      <Field label="Make 웹훅 URL (상담)" desc="상담 접수/확정/취소/리마인더/리뷰 등">
        <input value={webhookConsultation} onChange={(e) => setWebhookConsultation(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      <Field label="Make 웹훅 URL (AS접수)" desc="복원수리 접수 안내 (방문수거/직접발송/카운터)">
        <input value={webhookAsReceived} onChange={(e) => setWebhookAsReceived(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      <Field label="Make 웹훅 URL (AS상태변경)" desc="입고확인/입금안내/출고&송장/취소/만족도">
        <input value={webhookRepair} onChange={(e) => setWebhookRepair(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      <Field label="Make 웹훅 URL (이벤트)" desc="EVENT 접수확인/입금확인/출고완료 — 비우면 상담 웹훅으로 폴백">
        <input value={webhookEvent} onChange={(e) => setWebhookEvent(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      <Field label="Make 웹훅 URL (아임웹 주문)" desc="아임웹 주문 취소 접수·완료 / 반품 접수·승인·완료 — 비워 두면 발송하지 않음(폴백 없음). 솔라피 승인 + Make 분기 완성 후 입력 = 가동">
        <input value={webhookImweb} onChange={(e) => setWebhookImweb(e.target.value)}
          className="w-full h-9 px-3 rounded-lg border border-neutral-200 text-sm font-mono text-xs" placeholder="https://hook.eu2.make.com/..." />
      </Field>

      {/* 11-12. API 키 (읽기전용) */}
      <Field label="아임웹 API 키" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3 font-mono">
          {maskEnv('IMWEB_API_KEY')} — Vercel 환경변수에서 변경
        </p>
      </Field>

      <Field label="롯데택배 API 키" desc="">
        <p className="text-xs text-neutral-500 bg-neutral-50 rounded-lg p-3 font-mono">
          {maskEnv('LOTTE_CLIENT_KEY')} — Vercel 환경변수에서 변경
        </p>
      </Field>

      <div className="pt-4 border-t border-neutral-100">
        <Button onClick={handleSave} disabled={saving}><Save size={14} />{saving ? '저장 중...' : '저장'}</Button>
      </div>
    </div>
  );
}

function maskEnv(name: string) {
  return `${name}=****` + ' (보안상 마스킹)';
}

/** 푸시 테스트 패널 — 각 알림 타입별 테스트 발송 */
function PushTestPanel() {
  const [busy, setBusy] = useState<string | null>(null);

  const TESTS: Array<{ type: string; label: string; icon: string }> = [
    { type: 'generic',          label: '기본 테스트',                  icon: '🔔' },
    { type: 'review',           label: '리뷰 작성',                    icon: '⭐' },
    { type: 'consultation',     label: '상담 접수 (매장방문)',         icon: '📋' },
    { type: 'field_request',    label: '상담 접수 (출장)',             icon: '🚗' },
    { type: 'talk_received',    label: '상담 접수 (톡)',               icon: '💬' },
    { type: 'field_confirmed',  label: '출장 일정 확정',               icon: '✅' },
    { type: 'field_reschedule', label: '출장 일정 재요청',             icon: '🔄' },
    { type: 'repair_received',  label: '복원수리 접수',                icon: '🛠' },
    { type: 'order_received',   label: '아임웹 주문 접수',             icon: '📦' },
  ];

  const fire = async (type: string, label: string) => {
    setBusy(type);
    try {
      const res = await fetch('/api/push/test', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ type }),
      });
      const json = await res.json();
      if (!res.ok || !json.ok) {
        toast.error(`테스트 실패: ${json.error || 'unknown'}`);
        return;
      }
      const { sent, failed } = json.data;
      if (sent === 0 && failed === 0) {
        toast(`${label} 테스트 — 등록된 디바이스 없음 (알림 권한 허용 후 새로고침)`, { icon: '⚠️' });
      } else if (sent > 0) {
        toast.success(`${label} 테스트 발송 (${sent}건 성공${failed > 0 ? `, ${failed}건 실패` : ''})`);
      } else {
        toast.error(`${label} 테스트 — ${failed}건 모두 실패 (Vercel 로그 확인)`);
      }
    } catch (err) {
      toast.error(`테스트 실패: ${String(err)}`);
    } finally {
      setBusy(null);
    }
  };

  return (
    <div className="mt-3 pt-3 border-t border-neutral-200 space-y-2">
      <div className="flex items-center gap-1.5">
        <Bell size={13} className="text-neutral-600" />
        <span className="text-xs font-bold text-neutral-700">테스트 발송</span>
        <span className="text-[10px] text-neutral-400 ml-1">※ 실제 등록된 기기로 발송됩니다</span>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-1.5">
        {TESTS.map((t) => (
          <button
            key={t.type}
            onClick={() => fire(t.type, t.label)}
            disabled={busy !== null}
            className="flex items-center gap-1.5 px-2.5 py-1.5 rounded border border-neutral-200 bg-white hover:bg-neutral-50 text-[11px] text-neutral-700 disabled:opacity-50 text-left"
          >
            <span>{t.icon}</span>
            <span className="flex-1 truncate">{t.label}</span>
            {busy === t.type ? (
              <span className="text-[9px] text-neutral-400">...</span>
            ) : (
              <Send size={10} className="text-neutral-400" />
            )}
          </button>
        ))}
      </div>
      <p className="text-[10px] text-neutral-400 leading-relaxed">
        💡 모든 테스트가 실제 등록된 기기로 무조건 발송됩니다 (기기 연결 확인용).
      </p>
    </div>
  );
}

/** 알림 받는 기기 패널 (2026-09-15 재작성)
 *
 *  전엔 "이 기기만 알림 받기" 버튼 하나뿐이라 **몇 대가 등록돼 있는지 볼 수 없었다.**
 *  게다가 구독 API가 사용자당 토큰 1개만 남겨서, PC에서 열면 모바일이 끊기고
 *  모바일에서 열면 PC가 끊겼다(사장님 신고 → 근본원인).
 *  이제 기기 단위로 등록되므로, 등록된 기기를 **보여주고** 개별 해제만 제공한다.
 */
function DevicesPanel() {
  const [devices, setDevices] = useState<Array<{ id: string; label: string; isCurrent: boolean; updatedAt: string }>>([]);
  const [loading, setLoading] = useState(true);
  const [busyId, setBusyId] = useState<string | null>(null);
  const [iosNotice, setIosNotice] = useState(false);

  const load = async () => {
    try {
      const { getDeviceId, needsIosPwaNotice } = await import('@/lib/firebase/device');
      setIosNotice(needsIosPwaNotice());
      const res = await fetch(`/api/push/devices?deviceId=${encodeURIComponent(getDeviceId())}`);
      const json = await res.json();
      setDevices(json.devices || []);
    } catch {
      setDevices([]);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const handleRemove = async (id: string, label: string) => {
    if (!window.confirm(`"${label}" 기기의 알림을 해제합니다.\n이 기기에서 TMS를 다시 열면 자동으로 재등록됩니다.`)) return;
    setBusyId(id);
    try {
      const res = await fetch(`/api/push/devices?id=${id}`, { method: 'DELETE' });
      if (!res.ok) { toast.error('해제 실패'); return; }
      toast.success('해제 완료');
      await load();
    } finally {
      setBusyId(null);
    }
  };

  return (
    <div className="mt-3 pt-3 border-t border-neutral-200 space-y-2">
      <div className="flex items-center justify-between">
        <div className="flex items-center gap-1.5">
          <Smartphone size={13} className="text-neutral-600" />
          <span className="text-xs font-bold text-neutral-700">알림 받는 기기</span>
        </div>
        <button type="button" onClick={load} className="text-[11px] text-neutral-400 hover:text-neutral-700">새로고침</button>
      </div>

      {loading ? (
        <p className="text-[11px] text-neutral-400">불러오는 중…</p>
      ) : devices.length === 0 ? (
        <p className="text-[11px] text-amber-600 leading-relaxed bg-amber-50 rounded-lg px-2.5 py-2">
          등록된 기기가 없습니다. 알림이 오지 않습니다.<br />
          브라우저 알림 권한을 <b>허용</b>한 뒤 이 페이지를 새로고침해주세요.
        </p>
      ) : (
        <ul className="space-y-1">
          {devices.map((d) => (
            <li key={d.id} className="flex items-center gap-2 px-2.5 py-2 rounded-lg border border-neutral-200 bg-white">
              <span className="text-xs text-neutral-700 flex-1 truncate">
                {d.label}
                {d.isCurrent && <span className="ml-1.5 text-[10px] font-bold text-green-600">지금 이 기기</span>}
              </span>
              <button
                type="button"
                onClick={() => handleRemove(d.id, d.label)}
                disabled={busyId === d.id}
                className="text-[11px] text-neutral-400 hover:text-red-600 disabled:opacity-50 shrink-0"
              >
                {busyId === d.id ? '해제 중…' : '해제'}
              </button>
            </li>
          ))}
        </ul>
      )}

      <p className="text-[10px] text-neutral-400 leading-relaxed">
        알림 받을 기기에서 각각 TMS를 한 번 열고 알림 권한을 허용하면 자동 등록됩니다.
        PC·휴대폰 모두 등록해두면 양쪽 모두 알림이 옵니다.
      </p>

      {iosNotice && (
        <p className="text-[11px] text-amber-700 leading-relaxed bg-amber-50 rounded-lg px-2.5 py-2">
          <b>아이폰은 홈 화면에 추가해야 알림이 옵니다.</b><br />
          사파리에서 공유 버튼 → <b>홈 화면에 추가</b> → 홈 화면 아이콘으로 TMS를 연 뒤 알림을 허용해주세요.
          (iOS 정책상 사파리 탭에서는 웹 푸시가 오지 않습니다)
        </p>
      )}
    </div>
  );
}

function Field({ label, desc, children }: { label: string; desc?: string; children: React.ReactNode }) {
  return (<div><label className="block text-sm font-semibold text-neutral-800 mb-1">{label}</label>{desc && <p className="text-xs text-neutral-400 mb-2">{desc}</p>}{children}</div>);
}

function Toggle({ checked, onChange }: { checked: boolean; onChange: (v: boolean) => void }) {
  return (
    <button onClick={() => onChange(!checked)}
      className={`relative w-11 h-6 rounded-full transition ${checked ? 'bg-neutral-900' : 'bg-neutral-200'}`}>
      <span className={`absolute top-0.5 left-0.5 w-5 h-5 rounded-full bg-white shadow transition-transform ${checked ? 'translate-x-5' : ''}`} />
    </button>
  );
}
